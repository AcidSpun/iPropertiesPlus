Module globals
    ' GLobal Variables and Objects for the iPropertiesPlus add-in

    ' Sets a global variable for the Inventro Application
    Public g_inventorApplication As Inventor.Application

    'Sets a global variable for the Excel Application
    Public g_ExcelApp As New Microsoft.Office.Interop.Excel.Application

    'Opens the specific Excel workbook that contains the data for the combo/text boxs into a global variable
    Public g_wbProperties = g_ExcelApp.Workbooks.Open("G:\ALLCAD\Engineering Documents\INVENTOR\Custom Add-Ins\iProperties+\Properties.xlsx")

End Module
