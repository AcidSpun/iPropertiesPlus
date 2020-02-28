Module globals
    Public g_inventorApplication As Inventor.Application
    Public g_ExcelApp As New Microsoft.Office.Interop.Excel.Application
    Public g_wbProperties = g_ExcelApp.Workbooks.Open("G:\ALLCAD\Engineering Documents\INVENTOR\Custom Add-Ins\iProperties+\Properties.xlsx")
End Module
