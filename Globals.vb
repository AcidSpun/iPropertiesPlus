Module globals
    ' GLobal Variables and Objects for the iPropertiesPlus add-in

    ' Sets a global variable for the Inventro Application
    Public g_inventorApplication As Inventor.Application

    'Creating Global Variables
    Public g_startedRow As Integer
    Public g_totalRowsNext As Integer
    Public g_totalRowType As Integer
    Public g_totalRowRawMaterial As Integer
    Public g_totalRowSPClass As Integer
    Public g_totalRowTitle As Integer

    'Creating the Global Arrays
    Public g_nextArray(1, 2) As String
    Public g_typeArray(1, 3) As String
    Public g_matArray(1, 2) As String
    Public g_spClassArray(1) As String
    Public g_titleArray(1) As String

End Module
