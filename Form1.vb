﻿Imports Inventor
Imports Microsoft.Office.Interop

Public Class Form1
    'Declaration Part
#Disable Warning IDE0044 ' Add readonly modifier
    Dim Excel = New Microsoft.Office.Interop.Excel.Application
    Dim wbProperties = Excel.Workbooks.open("c:\iProperties+\Properties.xlsx")
    Dim nextProcessWS = wbProperties.Sheets(1)
    Dim typeWS = wbProperties.Sheets(2)
    Dim rawMaterialWS = wbProperties.Sheets(3)
    Dim SPClassWS = wbProperties.Sheets(4)
    Dim titleWS = wbProperties.Sheets(5)
#Enable Warning IDE0044 ' Add readonly modifier

#Disable Warning IDE1006 ' Naming Styles
    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click
#Enable Warning IDE1006 ' Naming Styles
        'Clean up Excel Workbooks
        releaseObject(wbProperties)
        releaseObject(nextProcessWS)
        releaseObject(typeWS)
        releaseObject(rawMaterialWS)
        releaseObject(SPClassWS)

        'Close the workbook
        wbProperties.Close()

        'Close Excel
        Excel.Quit()

        'Clean up Excel Object
        releaseObject(Excel)

        'Close Program
        Me.Close()
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'Declaration Part
        Dim startedRow As Integer
        Dim totalRowsNext As Integer
        Dim totalRowType As Integer
        Dim totalRowRawMaterial As Integer
        Dim totalRowSPClass As Integer
        Dim totalRowTitle As Integer

        'Clear Data from Comboboxes
        Me.cbNextProcess.Items.Clear()
        Me.cbRawMaterial.Items.Clear()
        Me.cbSPClass.Items.Clear()
        Me.cbType.Items.Clear()
        Me.cbSPClass.Items.Clear()
        Me.cbTitle.Items.Clear()

        'count number of rows in worksheets
        totalRowsNext = nextProcessWS.range("a1").Currentregion.Rows.Count
        totalRowType = typeWS.range("a1").Currentregion.Rows.Count
        totalRowRawMaterial = rawMaterialWS.range("a1").Currentregion.Rows.Count
        totalRowSPClass = SPClassWS.range("a1").Currentregion.Rows.Count
        totalRowTitle = titleWS.range("a1").Currentregion.Rows.Count

        'Loops for Populating the Combo Boxes
        For startedRow = 1 To totalRowsNext
            Me.cbNextProcess.Items.Add(nextProcessWS.Cells(startedRow, 1).text)
        Next

        For startedRow = 1 To totalRowType
            Me.cbType.Items.Add(typeWS.Cells(startedRow, 1).text)
        Next

        For startedRow = 1 To totalRowRawMaterial
            Me.cbRawMaterial.Items.Add(rawMaterialWS.Cells(startedRow, 2).text)
        Next

        For startedRow = 1 To totalRowSPClass
            Me.cbSPClass.Items.Add(SPClassWS.Cells(startedRow, 1).text)
        Next

        For startedRow = 1 To totalRowTitle
            Me.cbTitle.Items.Add(titleWS.Cells(startedRow, 1).text)
        Next

        readiProperty()

    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub releaseObject(ByVal obj As Object) 'Clean up Sub
#Enable Warning IDE1006 ' Naming Styles
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Private Sub cbNextProcess_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbNextProcess.SelectedIndexChanged
#Enable Warning IDE1006 ' Naming Styles

        'Decleration Part
        Dim startedRow As Integer
        Dim totalRows As Integer

        'count number of rows in worksheet
        totalRows = Excel.ActiveWorkbook.Sheets(1).Range("a1").Currentregion.Rows.Count

        'Add the Next Process Key to the corresponding text box
        For startedRow = 1 To totalRows
            If Me.cbNextProcess.Text = nextProcessWS.Cells(startedRow, 1).text Then
                Me.tbNextProcessKey.Text = nextProcessWS.Cells(startedRow, 2).text
            End If
        Next

    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Private Sub cbRawMaterial_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbRawMaterial.SelectedIndexChanged
#Enable Warning IDE1006 ' Naming Styles

        'Decleration Part
        Dim startedRow As Integer
        Dim totalRows As Integer

        'count number of rows in worksheet
        totalRows = Excel.ActiveWorkbook.Sheets(3).Range("a1").Currentregion.Rows.Count

        'Add the Raw Material Part Number to the corresponding text box
        For startedRow = 1 To totalRows
            If Me.cbRawMaterial.Text = rawMaterialWS.Cells(startedRow, 2).text Then
                Me.tbRawMaterialPartNumber.Text = rawMaterialWS.Cells(startedRow, 1).text
            End If
        Next

    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Private Sub cbType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbType.SelectedIndexChanged
#Enable Warning IDE1006 ' Naming Styles

        'Decleration Part
        Dim startedRow As Integer
        Dim totalRows As Integer

        'count number of rows in worksheet
        totalRows = Excel.ActiveWorkbook.Sheets(2).Range("a1").Currentregion.Rows.Count

        'Add the Type and Propertie to the corresponding text boxs
        For startedRow = 1 To totalRows
            If Me.cbType.Text = typeWS.Cells(startedRow, 1).text Then
                Me.tbTypeNumber.Text = typeWS.Cells(startedRow, 2).text
                Me.tbPropertyType.Text = typeWS.Cells(startedRow, 3).text
            End If
        Next

    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Private Sub btOK_Click(sender As Object, e As EventArgs) Handles btOK.Click
#Enable Warning IDE1006 ' Naming Styles

        'Declaration part
        Dim oApp As Inventor.Application
        Dim oDoc As Document
        Dim oPropSets As PropertySets
        Dim oPropSet As PropertySet
        Dim oDescription As [Property]
        Dim over As Integer

        'Get the active Document
        oApp = GetObject(, "Inventor.Application")
        oDoc = oApp.ActiveDocument

        'Get the PropertySets object
        oPropSets = oDoc.PropertySets

        'get the design tracking property set
        oPropSet = oPropSets.Item("Design Tracking Properties")

        'Get the Description iProperty
        oDescription = oPropSet.Item("Description")

        'Get the new description from the text box
        oDescription.Value = tbDescription.Text

        If Len(oDescription.Value) > 60 Then
            over = Len(oDescription.Value) - 60
            MsgBox("The Description may only have 60 Charecters." & vbCrLf & "Remove " & over & " Charecters")

        Else
            writeiProperty()

            'Clean up Excel Workbooks
            releaseObject(wbProperties)
            releaseObject(nextProcessWS)
            releaseObject(typeWS)
            releaseObject(rawMaterialWS)
            releaseObject(SPClassWS)

            'Close the workbook
            wbProperties.Close()

            'Close Excel
            Excel.Quit()

            'Clean up Excel Object
            releaseObject(Excel)

            'Close Program
            Me.Close()
        End If


    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Public Sub readiProperty()
#Enable Warning IDE1006 ' Naming Styles

        'Declaration part
        Dim oApp As Inventor.Application
        Dim oDoc As Document
        Dim oPropSets As PropertySets
        Dim oPropSet As PropertySet
        Dim oTitle As [Property]
        Dim oDescription As [Property]
        Dim oDefault As [Property]
        Dim oType As [Property]
        Dim oProp As [Property]
        Dim oMaterial As [Property]
        Dim oMaterialNum As [Property]
        Dim oNextProcess As [Property]
        Dim oNextProcessKey As [Property]
        Dim oSPClass As [Property]
        Dim oManufacturer As [Property]
        Dim oManPartNum As [Property]
        Dim oPropExists As Boolean = True

        oApp = GetObject(, "Inventor.Application")

        'Get the active Document
        oDoc = oApp.ActiveDocument

        'Get the PropertySets object
        oPropSets = oDoc.PropertySets

        'Get the summary property set
        oPropSet = oPropSets.Item("Inventor Summary Information")

        'Get the Title iProperty
        oTitle = oPropSet.Item("Title")

        'get the design tracking property set
        oPropSet = oPropSets.Item("Design Tracking Properties")

        'Get the Description iProperty
        oDescription = oPropSet.Item("Description")

        'Change the design tracking property set to custom
        oPropSet = oPropSets.Item("Inventor User Defined Properties")

        'get the custom design tracking properties

        'Get the custom design tracking properties if they exist and create them if they do not
        'Get or create Default Unit Property
        Try
            oDefault = oPropSet.Item("DEFAULT UNIT")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
#Disable Warning IDE0059 ' Unnecessary assignment of a value
            oDefault = oPropSet.Add("NUMBER", "DEFAULT UNIT")
#Enable Warning IDE0059 ' Unnecessary assignment of a value
            oPropExists = True
        End If

        'get or create the Type Property
        Try
            oType = oPropSet.Item("TYPE")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oType = oPropSet.Add("", "TYPE")
            oPropExists = True
        End If

        'get or create the Property property
        Try
            oProp = oPropSet.Item("PROPERTY")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oPropSet.Add("", "PROPERTY")
            oProp = oPropSet.Item("PROPERTY")
            oPropExists = True
        End If

        'Get or create the Material property
        Try
            oMaterial = oPropSet.Item("MATERIAL")
        Catch es As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oPropSet.Add("Select Raw Material", "MATERIAL")
            oMaterial = oPropSet.Item("MATERIAL")
            oPropExists = True
        End If

        'Get or create the Raw material Part Number property
        Try
            oMaterialNum = oPropSet.Item("RAW MATERIAL PART NUMBER")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oPropSet.Add("", "RAW MATERIAL PART NUMBER")
            oMaterialNum = oPropSet.Item("RAW MATERIAL PART NUMBER")
            oPropExists = True
        End If

        'get or create the Next Process property
        Try
            oNextProcess = oPropSet.Item("NEXT PROCESS")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oPropSet.Add("Select Next Process", "NEXT PROCESS")
            oNextProcess = oPropSet.Item("NEXT PROCESS")
            oPropExists = True
        End If

        'get or create the Next Process Key property
        Try
            oNextProcessKey = oPropSet.Item("NEXT PROCESS KEY")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oPropSet.Add("", "NEXT PROCESS KEY")
            oNextProcessKey = oPropSet.Item("NEXT PROCESS KEY")
            oPropExists = True
        End If

        'Get or create the SP Class property
        Try
            oSPClass = oPropSet.Item("SP CLASS")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oPropSet.Add("Select SP Class", "SP CLASS")
            oSPClass = oPropSet.Item("SP CLASS")
            oPropExists = True
        End If

        'Get or create the Manufaturer property
        Try
            oManufacturer = oPropSet.Item("MANUFACTURER")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oPropSet.Add("", "MANUFACTURER")
            oManufacturer = oPropSet.Item("MANUFACTURER")
            oPropExists = True
        End If

        'Get or create the Manufaturer Part Number property
        Try
            oManPartNum = oPropSet.Item("MANUFACTURER PART NUMBER")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oPropSet.Add("", "MANUFACTURER PART NUMBER")
            oManPartNum = oPropSet.Item("MANUFACTURER PART NUMBER")
        End If

#Disable Warning BC42104 ' Variable is used before it has been assigned a value
        'read the inventor iproperties
        cbTitle.Text = oTitle.Value
        tbDescription.Text = oDescription.Value
        tbTypeNumber.Text = oType.Value
        tbPropertyType.Text = oProp.Value
        cbRawMaterial.Text = oMaterial.Value
        tbRawMaterialPartNumber.Text = oMaterialNum.Value
        cbNextProcess.Text = oNextProcess.Value
        tbNextProcessKey.Text = oNextProcessKey.Value
        cbSPClass.Text = oSPClass.Value
        tbManufacturer.Text = oManufacturer.Value
        tbManPartNum.Text = oManPartNum.Value
#Enable Warning BC42104 ' Variable is used before it has been assigned a value

    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Public Sub writeiProperty()
#Enable Warning IDE1006 ' Naming Styles

        'Declaration part
        Dim oApp As Inventor.Application
        Dim oDoc As Document
        Dim oPropSets As PropertySets
        Dim oPropSet As PropertySet
        Dim oTitle As [Property]
        Dim oDescription As [Property]
        Dim oDefault As [Property]
        Dim oType As [Property]
        Dim oProp As [Property]
        Dim oMaterial As [Property]
        Dim oMaterialNum As [Property]
        Dim oNextProcess As [Property]
        Dim oNextProcessKey As [Property]
        Dim oSPClass As [Property]
        Dim oManufacturer As [Property]
        Dim oManPartNum As [Property]
        Dim oDelProp As [Property]
        Dim oPropCheck As Boolean = False

        'Get the active Document
        oApp = GetObject(, "Inventor.Application")
        oDoc = oApp.ActiveDocument

        'Get the PropertySets object
        oPropSets = oDoc.PropertySets

        'Get the summary property set
        oPropSet = oPropSets.Item("Inventor Summary Information")

        'Get the Title iProperty
        oTitle = oPropSet.Item("Title")

        'get the design tracking property set
        oPropSet = oPropSets.Item("Design Tracking Properties")

        'Get the Description iProperty
        oDescription = oPropSet.Item("Description")

        'Set the New description form the text Box
        oDescription.Value = tbDescription.Text

        'Change the design tracking property set to custom
        oPropSet = oPropSets.Item("Inventor User Defined Properties")

        'Get the custom design tracking properties
        oDefault = oPropSet("DEFAULT UNIT")
        oType = oPropSet.Item("TYPE")
        oProp = oPropSet.Item("PROPERTY")
        oMaterial = oPropSet.Item("MATERIAL")
        oMaterialNum = oPropSet.Item("RAW MATERIAL PART NUMBER")
        oNextProcess = oPropSet.Item("NEXT PROCESS")
        oNextProcessKey = oPropSet.Item("NEXT PROCESS KEY")
        oSPClass = oPropSet.Item("SP CLASS")
        oManufacturer = oPropSet.Item("MANUFACTURER")
        oManPartNum = oPropSet.Item("MANUFACTURER PART NUMBER")

        'Set the custom design tracking properties
        oDefault.Value = "Number"
        oTitle.Value = cbTitle.Text
        oType.Value = tbTypeNumber.Text
        oProp.Value = tbPropertyType.Text
        oMaterial.Value = cbRawMaterial.Text
        oMaterialNum.Value = tbRawMaterialPartNumber.Text
        oNextProcess.Value = cbNextProcess.Text
        oNextProcessKey.Value = tbNextProcessKey.Text
        oSPClass.Value = cbSPClass.Text
        oManufacturer.Value = tbManufacturer.Text
        oManPartNum.Value = tbManPartNum.Text

        'Check to see if the Italian properties exis, and delete them if they do
        Try
            oDelProp = oPropSet("it.system_group.TIPO")
        Catch ex As Exception
            oPropCheck = True
        End Try

        If Not oPropCheck Then
#Disable Warning BC42104 ' Variable is used before it has been assigned a value
            oDelProp.Delete()
#Enable Warning BC42104 ' Variable is used before it has been assigned a value
            oDelProp = oPropSet("it.system_group.PROPRIETA")
            oDelProp.Delete()
        End If

    End Sub

End Class