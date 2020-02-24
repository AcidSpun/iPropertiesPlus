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
        Private WithEvents m_ipropertyPlusButton As ButtonDefinition
        Private WithEvents m_UIEvents As UserInterfaceEvents
        Private WithEvents M_appEvents As ApplicationEvents

#Region "ApplicationAddInServer Members"

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' This method is called by Inventor when the AddIn is unloaded.
            ' The AddIn will be unloaded either manually by the user or
            ' when the Inventor session is terminated.

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            Marshal.ReleaseComObject(g_inventorApplication)
            g_inventorApplication = Nothing

            System.GC.WaitForPendingFinalizers()
            System.GC.Collect()

        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation

            ' This property is provided to allow the AddIn to expose an API 
            ' of its own to other programs. Typically, this  would be done by
            ' implementing the AddIn's API interface in a class and returning 
            ' that class object through this property.

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
            'Get a reference to the UserInterfaceManager object
            Dim UIManager As Inventor.UserInterfaceManager
            UIManager = g_inventorApplication.UserInterfaceManager

            'Add the command to the File Controls, just before the standard iProperties command
            Dim fileControls As CommandControls = UIManager.FileBrowserControls
            fileControls.AddButton(m_ipropertyPlusButton, , , "AppiPropertiesWrapperCmd", True)
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
                            fileCommandBar.Controls.AddButton(m_ipropertyPlusButton, cbControl.Index)

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
        Public Shared Sub Register(ByVal t As Type)

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
        Public Shared Sub Unregister(ByVal t As Type)

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
                Dim guid As String = ""
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

        Private Sub m_appEvents_OnActivateView(ByVal ViewObject As Inventor.View, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles M_appEvents.OnActivateView
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If g_inventorApplication.Views.Count > 0 Then
                    m_ipropertyPlusButton.Enabled = True
                End If
            End If
        End Sub

        Private Sub m_appEvents_OnDeactivateView(ByVal ViewObject As Inventor.View, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles M_appEvents.OnDeactivateView
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If g_inventorApplication.Views.Count > 0 Then
                    m_ipropertyPlusButton.Enabled = True
                Else
                    m_ipropertyPlusButton.Enabled = False
                End If
            End If
        End Sub

        Private Sub m_iPropertyPlusButton_OnExecute(ByVal Context As Inventor.NameValueMap) Handles m_ipropertyPlusButton.OnExecute
            Dim dialog As Form1
            dialog.ShowDialog()
        End Sub

        Public Sub Activate(AddInSiteObject As ApplicationAddInSite, FirstTime As Boolean) Implements ApplicationAddInServer.Activate
            Throw New NotImplementedException()
        End Sub
    End Class
End Namespace


