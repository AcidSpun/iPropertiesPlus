Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Namespace iPropertiesPlus
    <ProgIdAttribute("iPropertiesPlus.StandardAddInServer"),
    GuidAttribute("98bb4777-41d2-47ac-82c9-f56f4f3fe154")>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        ' Inventor application object.
        Private m_clientID As String
        Private WithEvents m_iPropertyPlusButton As ButtonDefinition
        Private WithEvents m_UIEvents As UserInterfaceEvents
        Private WithEvents m_appEvents As ApplicationEvents

#Region "ApplicationAddInServer Members"

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application
            m_UIEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents
            m_appEvents = g_inventorApplication.ApplicationEvents

            ' Set the member variable for the client ID.
            m_clientID = AddInGuid(Me.GetType)

            ' Get the icon for the button as an iPictureDisp object
            Dim buttonIcon As stdole.IPictureDisp = Microsoft.VisualBasic.Compatibility.VB6.IconToIPicture(My.Resources.iPropPlus)

            ' Create the button for the iProperty Plus command.
            m_iPropertyPlusButton = g_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition("iProperties +", "iPropertyPlus", CommandTypesEnum.kFilePropertyEditCmdType, m_clientID, "Custom iProperty command.", "iProperty +", buttonIcon, buttonIcon)

            ' Set the enabled state based on whether there are any visible documents or not.
            If g_inventorApplication.Views.Count > 0 Then
                m_iPropertyPlusButton.Enabled = True
            Else
                m_iPropertyPlusButton.Enabled = False
            End If

            If firstTime Then
                If g_inventorApplication.UserInterfaceManager.InterfaceStyle = InterfaceStyleEnum.kRibbonInterface Then
                    CreateOrUpdateRibbon()
                Else
                    CreateOrUpdateClassic()
                End If
            End If
        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
            ' Release objects.
            Marshal.ReleaseComObject(g_inventorApplication)
            g_inventorApplication = Nothing

            System.GC.WaitForPendingFinalizers()
            System.GC.Collect()
        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
            ' Note:this method is now obsolete, you should use the 
            ' ControlDefinition functionality for implementing commands.
        End Sub

#End Region

        Private Sub CreateOrUpdateRibbon()
            ' Get a reference to the UserInterfaceManager object.
            Dim UIManager As Inventor.UserInterfaceManager
            UIManager = g_inventorApplication.UserInterfaceManager

            ' Add the command to the File controls, just before the standard iProperties command. 
            Dim fileControls As CommandControls = UIManager.FileBrowserControls
            fileControls.AddButton(m_iPropertyPlusButton, , , "AppiPropertiesWrapperCmd", True)
        End Sub

        Private Sub CreateOrUpdateClassic()
            ' Add a button to the command bar that's used for the File menus of each of the environments.
            For Each currentEnvironment As Inventor.Environment In g_inventorApplication.UserInterfaceManager.Environments
                If Not currentEnvironment.DefaultMenuBar Is Nothing Then
                    Dim menuBar As Inventor.CommandBar = currentEnvironment.DefaultMenuBar

                    ' Get the command bar for the File menu, assuming it is the first one.
                    Dim fileCommandBar As Inventor.CommandBar = menuBar.Controls.Item(1).CommandBar

                    ' Find the standard iProperties command.
                    For Each cbControl As Inventor.CommandBarControl In fileCommandBar.Controls
                        ' Check to see if the iProperty+ command is already in the menu since some menus
                        ' are shared by multiple environments it may have already been added.
                        If cbControl.InternalName = "iPropertyPlus" Then
                            Exit For
                        ElseIf cbControl.InternalName = "AppiPropertiesWrapperCmd" Then
                            ' Add the custom button just before this one.
                            fileCommandBar.Controls.AddButton(m_iPropertyPlusButton, cbControl.Index)

                            ' Exit the loop since a match was found.
                            Exit For
                        End If
                    Next
                End If
            Next
        End Sub

        Private Sub m_UIEvents_OnResetCommandBars(ByVal CommandBars As Inventor.ObjectsEnumerator, ByVal Context As Inventor.NameValueMap) Handles m_UIEvents.OnResetCommandBars
            CreateOrUpdateClassic()
        End Sub

        Private Sub m_UIEvents_OnResetRibbonInterface(ByVal Context As Inventor.NameValueMap) Handles m_UIEvents.OnResetRibbonInterface
            CreateOrUpdateRibbon()
        End Sub

#Region "COM Registration"

        ' Registers this class as an AddIn for Inventor.
        ' This function is called when the assembly is registered for COM.
        <ComRegisterFunctionAttribute()>
        Private Shared Sub Register(ByVal t As Type)

            Dim clssRoot As RegistryKey = Registry.ClassesRoot
            Dim clsid As RegistryKey = Nothing
            Dim subKey As RegistryKey = Nothing

            Try
                clsid = clssRoot.CreateSubKey("CLSID\" + AddInGuid(t))
                clsid.SetValue(Nothing, "iPropertiesPlus")
                subKey = clsid.CreateSubKey("Implemented Categories\{39AD2B5C-7A29-11D6-8E0A-0010B541CAA8}")
                subKey.Close()

                subKey = clsid.CreateSubKey("Settings")
                subKey.SetValue("AddInType", "Standard")
                subKey.SetValue("LoadOnStartUp", "1")

                'subKey.SetValue("SupportedSoftwareVersionLessThan", "")
                subKey.SetValue("SupportedSoftwareVersionGreaterThan", "13..")
                'subKey.SetValue("SupportedSoftwareVersionEqualTo", "")
                'subKey.SetValue("SupportedSoftwareVersionNotEqualTo", "")
                'subKey.SetValue("Hidden", "0")
                'subKey.SetValue("UserUnloadable", "1")
                subKey.SetValue("Version", 0)
                subKey.Close()

                subKey = clsid.CreateSubKey("Description")
                subKey.SetValue(Nothing, "iPropertiesPlus")

            Catch ex As Exception
                System.Diagnostics.Trace.Assert(False)
            Finally
                If Not subKey Is Nothing Then subKey.Close()
                If Not clsid Is Nothing Then clsid.Close()
                If Not clssRoot Is Nothing Then clssRoot.Close()
            End Try

        End Sub

        ' Unregisters this class as an AddIn for Inventor.
        ' This function is called when the assembly is unregistered.
        <ComUnregisterFunctionAttribute()>
        Private Shared Sub Unregister(ByVal t As Type)

            Dim clssRoot As RegistryKey = Registry.ClassesRoot
            Dim clsid As RegistryKey = Nothing

            Try
                clssRoot = Microsoft.Win32.Registry.ClassesRoot
                clsid = clssRoot.OpenSubKey("CLSID\" + AddInGuid(t), True)
                clsid.SetValue(Nothing, "")
                clsid.DeleteSubKeyTree("Implemented Categories\{39AD2B5C-7A29-11D6-8E0A-0010B541CAA8}")
                clsid.DeleteSubKeyTree("Settings")
                clsid.DeleteSubKeyTree("Description")
            Catch
            Finally
                If Not clsid Is Nothing Then clsid.Close()
                If Not clssRoot Is Nothing Then clssRoot.Close()
            End Try

        End Sub

        ' This property uses reflection to get the value for the GuidAttribute attached to the class.
        Public Shared ReadOnly Property AddInGuid(ByVal t As Type) As String
            Get
                Dim guid As String = "98bb4777-41d2-47ac-82c9-f56f4f3fe154"
                Try
                    Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
                    Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
                    guid = "{" + guidAttribute.Value.ToString() + "}"
                Finally
                    AddInGuid = guid
                End Try
            End Get
        End Property

#End Region

        Private Sub m_appEvents_OnActivateView(ByVal ViewObject As Inventor.View, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_appEvents.OnActivateView
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If g_inventorApplication.Views.Count > 0 Then
                    m_iPropertyPlusButton.Enabled = True
                End If
            End If
        End Sub

        Private Sub m_appEvents_OnDeactivateView(ByVal ViewObject As Inventor.View, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_appEvents.OnDeactivateView
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If g_inventorApplication.Views.Count > 0 Then
                    m_iPropertyPlusButton.Enabled = True
                Else
                    m_iPropertyPlusButton.Enabled = False
                End If
            End If
        End Sub

        Private Sub m_iPropertyPlusButton_OnExecute(ByVal Context As Inventor.NameValueMap) Handles m_iPropertyPlusButton.OnExecute
            Using dialog As New fmiPropertiesPlus
                dialog.ShowDialog()
            End Using
        End Sub
    End Class

End Namespace