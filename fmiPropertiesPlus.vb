Imports System.Runtime.InteropServices
Imports Inventor

Public Class fmiPropertiesPlus

    'Declaration Part for Class Global Variables
    Public ReadOnly nextProcessWS As Object = g_wbProperties.Sheets(1)   'Gets the Next Process Worksheet from Excel
    Public ReadOnly typeWS As Object = g_wbProperties.Sheets(2)          'Gets the Type Worksheet from Excel
    Public ReadOnly rawMaterialWS As Object = g_wbProperties.Sheets(3)   'Gets the Raw Materials Worksheet from Excel
    Public ReadOnly SPClassWS As Object = g_wbProperties.Sheets(4)       'Gets the SP Class Worksheet from Excel
    Public ReadOnly titleWS As Object = g_wbProperties.Sheets(5)         'Gets the Title Worksheet from Excel


    Private Sub fmiProperteisPlus_Shown(sender As Object, e As EventArgs) Handles Me.Shown 'This Sub runs when the iProperties+ window is shown

        cbBoxFill() 'Populates the combo boxes from the excel worksheet
        readiProperty() ' reads the current iProperties from the inventor files, and populates the related combo/text boxes

    End Sub

    Private Sub cbBoxFill() 'Populates the Combo Boxes from the Excel Spreadsheet Data

        'Declaration Part for local variables
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


#Region "Populate Combo Box Menues"

        'Loops for Populating the Combo Boxes

        ' Populates the Next Process drop down menu
        For startedRow = 1 To totalRowsNext
            Me.cbNextProcess.Items.Add(nextProcessWS.Cells(startedRow, 1).text)
        Next

        ' Populates the Type drop down menu
        For startedRow = 1 To totalRowType
            Me.cbType.Items.Add(typeWS.Cells(startedRow, 1).text)
        Next

        ' Populates the Raw Material drop down menu
        For startedRow = 1 To totalRowRawMaterial
            Me.cbRawMaterial.Items.Add(rawMaterialWS.Cells(startedRow, 2).text)
        Next

        ' Populates the SP Class drop down menu
        For startedRow = 1 To totalRowSPClass
            Me.cbSPClass.Items.Add(SPClassWS.Cells(startedRow, 1).text)
        Next

        ' Populates the Title drop down menu
        For startedRow = 1 To totalRowTitle
            Me.cbTitle.Items.Add(titleWS.Cells(startedRow, 1).text)
        Next

#End Region

    End Sub

    Public Sub readiProperty()

        'Declaration part for local variables
        Dim oApp As Inventor.Application
        Dim oDoc As Document
        Dim oPropSets As PropertySets
        Dim oPropSet As PropertySet
        Dim oTitle As [Property]
        Dim oDescription As [Property]
        Dim oDefault As [Property]
        Dim oTypeName As [Property]
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

        'Get the Inventor.Application Object
        oApp = GetObject(, "Inventor.Application")

        'Get the active Document
        oDoc = oApp.ActiveDocument

        'Get the PropertySets object
        oPropSets = oDoc.PropertySets

        'Get the Summary property set from Inventor files iProperties to get Title
        oPropSet = oPropSets.Item("Inventor Summary Information")

        'Get the Title iProperty
        oTitle = oPropSet.Item("Title")

        'change the property set to Design Tracking Properties to get Description
        oPropSet = oPropSets.Item("Design Tracking Properties")

        'Get the Description iProperty
        oDescription = oPropSet.Item("Description")

        'Change the design tracking property set to Custom to get the required properties for WindChill use
        oPropSet = oPropSets.Item("Inventor User Defined Properties")

        'get the custom design tracking properties

#Region "Get or Create the Custom Design Tracking Properties"


        'Get the custom design tracking properties if they exist and create them if they do not
        'Get or create Default Unit Property
        Try
            oDefault = oPropSet.Item("DEFAULT UNIT")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oDefault = oPropSet.Add("NUMBER", "DEFAULT UNIT")
            oDefault = oPropSet.Item("DEFAULT UNIT")
            oPropExists = True
        End If

        'get or create the Type Name Property
        Try
            oTypeName = oPropSet.Item("TYPE NAME")
        Catch ex As Exception
            oPropExists = False
        End Try

        If Not oPropExists Then
            oTypeName = oPropSet.Add("Select Type", "TYPE NAME")
            oTypeName = oPropSet.Item("TYPE NAME")
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
            oType = oPropSet.Item("TYPE")
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

#End Region


        ' Populate the Text/Combo boxes with the current iProperties values
        cbTitle.Text = oTitle.Value
        tbDescription.Text = oDescription.Value
        cbType.Text = oTypeName.Value
        tbTypeNumber.Text = oType.Value
        tbPropertyType.Text = oProp.Value
        cbRawMaterial.Text = oMaterial.Value
        tbRawMaterialPartNumber.Text = oMaterialNum.Value
        cbNextProcess.Text = oNextProcess.Value
        tbNextProcessKey.Text = oNextProcessKey.Value
        cbSPClass.Text = oSPClass.Value
        tbManufacturer.Text = oManufacturer.Value
        tbManPartNum.Text = oManPartNum.Value

    End Sub

    Private Sub cbNextProcess_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbNextProcess.SelectedIndexChanged
        'Populates text boxs related to the Next Process combo box on item select

        'Decleration Part for local variables
        Dim startedRow As Integer
        Dim totalRows As Integer

        'count number of rows inNext Process worksheet
        totalRows = g_ExcelApp.ActiveWorkbook.Sheets(1).Range("a1").Currentregion.Rows.Count

        'Add the Next Process Key to the corresponding text box
        For startedRow = 1 To totalRows
            If cbNextProcess.Text = nextProcessWS.Cells(startedRow, 1).text Then
                tbNextProcessKey.Text = nextProcessWS.Cells(startedRow, 2).text
            End If
        Next

    End Sub

    Private Sub cbRawMaterial_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbRawMaterial.SelectedIndexChanged
        'Populates text boxs related to the raw Materials combo box on item select

        'Decleration Part for local variables
        Dim startedRow As Integer
        Dim totalRows As Integer

        'count number of rows in Raw Material worksheet
        totalRows = g_ExcelApp.ActiveWorkbook.Sheets(3).Range("a1").Currentregion.Rows.Count

        'Add the Raw Material Part Number to the corresponding text box
        For startedRow = 1 To totalRows
            If cbRawMaterial.Text = rawMaterialWS.Cells(startedRow, 2).text Then
                tbRawMaterialPartNumber.Text = rawMaterialWS.Cells(startedRow, 1).text
            End If
        Next

    End Sub

    Private Sub cbType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbType.SelectedIndexChanged
        'Populates text boxs related to the Type combo box on item select

        'Decleration Part for local variables
        Dim startedRow As Integer
        Dim totalRows As Integer

        'count number of rows in Type worksheet
        totalRows = g_ExcelApp.ActiveWorkbook.Sheets(2).Range("a1").Currentregion.Rows.Count

        'Add the Type and Propertie to the corresponding text boxs
        For startedRow = 1 To totalRows
            If cbType.Text = typeWS.Cells(startedRow, 1).text Then
                tbTypeNumber.Text = typeWS.Cells(startedRow, 2).text
                tbPropertyType.Text = typeWS.Cells(startedRow, 3).text
            End If
        Next

    End Sub

    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click 'Cancel Button Clicked

        Close() 'Close the iProperties+ window

    End Sub

    Private Sub btOK_Click(sender As Object, e As EventArgs) Handles btOK.Click 'OK Button Clicked

        'Declaration part for local variables
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

        ' Check that the description is less than 60 charecters.  If over 60 then inform the user of how many charecters they are over by
        ' If under 60 charecters then write the new iproperties to inventor
        If Len(oDescription.Value) > 60 Then
            over = Len(oDescription.Value) - 60
            MsgBox("The Description may only have 60 Charecters." & vbCrLf & "Remove " & over & " Charecters")

        Else
            writeiProperty() 'Calls the sub to write the iProperties from the text/combo boxs
        End If

        Close() 'Close the iProperties+ window

    End Sub

    Public Sub writeiProperty()

        'Declaration part
        Dim oApp As Inventor.Application
        Dim oDoc As Document
        Dim oPropSets As PropertySets
        Dim oPropSet As PropertySet
        Dim oTitle As [Property]
        Dim oDescription As [Property]
        Dim oDefault As [Property]
        Dim oTypeName As [Property]
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
        oTypeName = oPropSet("TYPE NAME")
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
        oTypeName.Value = cbType.Text
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
            oDelProp.Delete()
            oDelProp = oPropSet("it.system_group.PROPRIETA")
            oDelProp.Delete()
        End If

    End Sub

    Private Sub fmiProperteisPlus_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        ' This is what happens when the iProperties+ window is closed

        ' Release the worksheets
        Marshal.ReleaseComObject(typeWS)
        Marshal.ReleaseComObject(nextProcessWS)
        Marshal.ReleaseComObject(rawMaterialWS)
        Marshal.ReleaseComObject(SPClassWS)
        Marshal.ReleaseComObject(titleWS)

        'Cleanup
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

End Class