Imports System.Runtime.InteropServices

Module readDataFromExcel
    Public Sub readDataToArrays() 'Reads the Excel Data into Arrays in system memory
        Try
            'Sets a variable for the Excel Application
            Dim ExcelApp As New Microsoft.Office.Interop.Excel.Application
            ExcelApp.Visible = False
            ExcelApp.UserControl = False

            'Opens the specific Excel workbook that contains the data for the combo/text boxs into variables
            Dim wbProperties As Object = ExcelApp.Workbooks.Open("G:\ALLCAD\Engineering Documents\INVENTOR\Custom Add-Ins\iProperties+\Properties.xlsx")
            Dim nextProcessWS As Object = wbProperties.Sheets(1)   'Gets the Next Process Worksheet from Excel
            Dim typeWS As Object = wbProperties.Sheets(2)          'Gets the Type Worksheet from Excel
            Dim rawMaterialWS As Object = wbProperties.Sheets(3)   'Gets the Raw Materials Worksheet from Excel
            Dim SPClassWS As Object = wbProperties.Sheets(4)       'Gets the SP Class Worksheet from Excel
            Dim titleWS As Object = wbProperties.Sheets(5)         'Gets the Title Worksheet from Excel

            'Count the number of rows in the worksheets
            g_totalRowsNext = nextProcessWS.range("a1").Currentregion.Rows.Count
            g_totalRowType = typeWS.range("a1").Currentregion.Rows.Count
            g_totalRowRawMaterial = rawMaterialWS.range("a1").Currentregion.Rows.Count
            g_totalRowSPClass = SPClassWS.range("a1").Currentregion.Rows.Count
            g_totalRowTitle = titleWS.range("a1").Currentregion.Rows.Count

            'Redimension Global Arrays
            ReDim g_nextArray(g_totalRowsNext - 1, 2)
            ReDim g_typeArray(g_totalRowType - 1, 3)
            ReDim g_matArray(g_totalRowRawMaterial - 1, 2)
            ReDim g_spClassArray(g_totalRowSPClass - 1)
            ReDim g_titleArray(g_totalRowTitle - 1)

            'Populating the Arrays from Excel
            For g_startedRow = 1 To g_totalRowsNext
                g_nextArray(g_startedRow - 1, 1) = nextProcessWS.Cells(g_startedRow, 1).text
                g_nextArray(g_startedRow - 1, 2) = nextProcessWS.Cells(g_startedRow, 2).text
            Next

            For g_startedRow = 1 To g_totalRowType
                g_typeArray(g_startedRow - 1, 1) = typeWS.Cells(g_startedRow, 1).text
                g_typeArray(g_startedRow - 1, 2) = typeWS.Cells(g_startedRow, 2).text
                g_typeArray(g_startedRow - 1, 3) = typeWS.Cells(g_startedRow, 3).text
            Next

            For g_startedRow = 1 To g_totalRowRawMaterial
                g_matArray(g_startedRow - 1, 1) = rawMaterialWS.Cells(g_startedRow, 1).text
                g_matArray(g_startedRow - 1, 2) = rawMaterialWS.Cells(g_startedRow, 2).text
            Next

            For g_startedRow = 1 To g_totalRowSPClass
                g_spClassArray(g_startedRow - 1) = SPClassWS.Cells(g_startedRow).text
            Next

            For g_startedRow = 1 To g_totalRowTitle
                g_titleArray(g_startedRow - 1) = titleWS.Cells(g_startedRow).text
            Next

            'Close the Sheets
            NAR(nextProcessWS)
            NAR(typeWS)
            NAR(rawMaterialWS)
            NAR(SPClassWS)
            NAR(titleWS)


            'Close the Workbook
            wbProperties.Close()
            NAR(wbProperties)

            'Close Excel
            ExcelApp.close()
            ExcelApp.Quit()
            NAR(ExcelApp)

            'Cleanup
            GC.WaitForPendingFinalizers()
            GC.Collect()

        Catch ex As Exception

            MsgBox(ex.Message.ToString)

        End Try

    End Sub

    Private Sub NAR(o As Object) 'Closes and cleans up the Excel COM Objects
        Try
            While Marshal.ReleaseComObject(o) > 0
            End While
        Catch ex As Exception
        Finally
            o = Nothing
        End Try
    End Sub
End Module
