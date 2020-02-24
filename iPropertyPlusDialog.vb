Imports System.Windows.Forms

Public Class iPropertyPlusDialog
    Private m_doc As Inventor.Document

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim trans As Inventor.Transaction = Nothing
        Try
            ' Start a transaction.
            trans = g_inventorApplication.TransactionManager.StartTransaction(m_doc, "iProperties+ Edit")

            ' Update the property values.
            m_doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = Me.txtPartNumber.Text
            m_doc.PropertySets.Item("Design Tracking Properties").Item("Standard Revision").Value = Me.txtRevisionNumber.Text
            m_doc.PropertySets.Item("Design Tracking Properties").Item("Description").Value = Me.txtDescription.Text
            m_doc.PropertySets.Item("Design Tracking Properties").Item("Designer").Value = Me.txtDesigner.Text
            m_doc.PropertySets.Item("Design Tracking Properties").Item("Engr Approved By").Value = Me.cboApprovedBy.Text

            ' Update the approved date.
            Dim noDate As Date = Date.FromFileTimeUtc(0)
            If Me.datApprovedDate.Checked Then
                m_doc.PropertySets.Item("Design Tracking Properties").Item("Engr Date Approved").Value = Me.datApprovedDate.Value
            Else
                m_doc.PropertySets.Item("Design Tracking Properties").Item("Engr Date Approved").Value = noDate
            End If

            ' Update the creation date.
            If Me.datCreationDate.Checked Then
                m_doc.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value = Me.datCreationDate.Value
            Else
                m_doc.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value = noDate
            End If

            Dim tempDate As Date = m_doc.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value
            If tempDate.ToFileTimeUtc <> 0 Then
                Me.datCreationDate.Checked = True
                Me.datCreationDate.Value = m_doc.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value
            End If

            ' Update the material, if it's different and this is a part.
            If Me.cboMaterial.Enabled Then
                Dim partDoc As Inventor.PartDocument = m_doc

                If Me.cboMaterial.Text <> partDoc.ComponentDefinition.Material.Name Then
                    partDoc.ComponentDefinition.Material = partDoc.Materials.Item(Me.cboMaterial.Text)
                End If
            End If

            ' Update the finish.  This is a custom property and will need to be created if it doesn't exist.
            Dim finishProperty As Inventor.Property = Nothing
            Try
                finishProperty = m_doc.PropertySets.Item("Inventor User Defined Properties").Item("Finish")
            Catch ex As Exception
            End Try

            If finishProperty Is Nothing Then
                ' Create the finish property.
                finishProperty = m_doc.PropertySets.Item("Inventor User Defined Properties").Add(cboFinish.Text, "Finish")
            Else
                finishProperty.Value = cboFinish.Text
            End If


        Catch ex As Exception
            MsgBox("Unexpected error updating the iProperties.  Aborting the operation.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)

            If Not trans Is Nothing Then
                trans.Abort()
            End If

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()

            Exit Sub
        End Try

        trans.End()

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub iPropertyPlusDialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ' Get the active edit document.
            m_doc = g_inventorApplication.ActiveEditDocument

            ' Initialize the values on the dialog using the current property values.
            Me.txtPartNumber.Text = m_doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
            Me.txtRevisionNumber.Text = m_doc.PropertySets.Item("Design Tracking Properties").Item("Standard Revision").Value
            Me.txtDescription.Text = m_doc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
            Me.txtDesigner.Text = m_doc.PropertySets.Item("Design Tracking Properties").Item("Designer").Value

            Dim tempDate As Date = m_doc.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value
            If tempDate.ToFileTimeUtc <> 0 Then
                Me.datCreationDate.Checked = True
                Me.datCreationDate.Value = m_doc.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value
            End If

            tempDate = m_doc.PropertySets.Item("Design Tracking Properties").Item("Engr Date Approved").Value
            If tempDate.ToFileTimeUtc <> 0 Then
                Me.datApprovedDate.Checked = True
                Me.datApprovedDate.Value = m_doc.PropertySets.Item("Design Tracking Properties").Item("Engr Date Approved").Value
            End If

            ' Populate the Materials list.
            If m_doc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
                Dim partDoc As Inventor.PartDocument = CType(m_doc, Inventor.PartDocument)

                For Each currentMaterial As Inventor.Material In partDoc.Materials
                    Me.cboMaterial.Items.Add(currentMaterial.Name)
                Next

                ' Select the current material from the list.
                Me.cboMaterial.SelectedItem = partDoc.ComponentDefinition.Material.Name
            Else
                Me.cboMaterial.Enabled = False
            End If

            ' Open the xml document that contains the lists.
            Dim listFilename As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location) & "\iPropertyPlusList.xml"
            Dim listxmlDoc As New System.Xml.XmlDocument
            listxmlDoc.Load(listFilename)

            ' Read the finish list from the xml file.
            Dim finishNode As System.Xml.XmlNodeList = listxmlDoc.GetElementsByTagName("Finish")
            For Each itemNode As System.Xml.XmlNode In finishNode.Item(0).ChildNodes
                Me.cboFinish.Items.Add(itemNode.InnerText)
            Next

            ' Select the value based on the current setting.
            Dim finishValue As String = ""
            Try
                finishValue = m_doc.PropertySets.Item("Inventor User Defined Properties").Item("Finish").Value
            Catch ex As Exception

            End Try

            If finishValue <> "" Then
                Me.cboFinish.SelectedItem = finishValue
            End If

            ' Read the approved by list from the xml file.
            Dim approvedByNode As System.Xml.XmlNodeList = listxmlDoc.GetElementsByTagName("ApprovedBy")
            For Each itemNode As System.Xml.XmlNode In approvedByNode.Item(0).ChildNodes
                Me.cboApprovedBy.Items.Add(itemNode.InnerText)
            Next

            ' Select the value based on the current setting.
            Dim approvedByValue As String = ""
            Try
                approvedByValue = m_doc.PropertySets.Item("Design Tracking Properties").Item("Engr Approved By").Value
            Catch ex As Exception

            End Try

            If approvedByValue <> "" Then
                Me.cboApprovedBy.SelectedItem = approvedByValue
            End If

        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub
End Class
