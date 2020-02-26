Module globals
    Public g_inventorApplication As Inventor.Application
    Public ReadOnly g_Excel = New Microsoft.Office.Interop.Excel.Application
    Public ReadOnly g_wbProperties = g_Excel.Workbooks.Open("G:\ALLCAD\Engineering Documents\INVENTOR\Custom Add-Ins\iProperties+\Properties.xlsx")
End Module