Imports System.Runtime.InteropServices


Namespace DoyleAddin
    <ProgIdAttribute("Test2.StandardAddInServer"),
    GuidAttribute("513b9d7e-103e-4569-8eb5-ab3929cd33ad")>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        Private WithEvents UiEvents As UserInterfaceEvents
        Private WithEvents DXFUpdate As ButtonDefinition
        Private WithEvents PrintUpdate As ButtonDefinition

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            ' Initialize AddIn members.
            ThisApplication = addInSiteObject.Application

            ' Get a reference to the UserInterfaceManager object. 
            Dim UIManager As Inventor.UserInterfaceManager = ThisApplication.UserInterfaceManager

            ' Get a reference to the ControlDefinitions object. 
            Dim controlDefs As ControlDefinitions = ThisApplication.CommandManager.ControlDefinitions

            ' TODO: Add button definitions.

            ' Sample to illustrate creating a button definition.
            Dim PrintUpdateIconLarge As stdole.IPictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.PrintUpdateIconLarge)
            Dim PrintUpdateIconSmall As stdole.IPictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.PrintUpdateIconSmall)
            Dim DXFUpdateIconSmall As stdole.IPictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.DXFUpdateIconSmall)
            Dim DXFUpdateIconLarge As stdole.IPictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.DXFUpdateIconLarge)
            DXFUpdate = controlDefs.AddButtonDefinition("DXF Update", "dxfUpdate", CommandTypesEnum.kShapeEditCmdType, AddInClientID, , , DXFUpdateIconSmall, DXFUpdateIconLarge)
            PrintUpdate = controlDefs.AddButtonDefinition("Print Update", "printUpdate", CommandTypesEnum.kShapeEditCmdType, AddInClientID, , , PrintUpdateIconSmall, PrintUpdateIconLarge)

            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If

            ' Connect to the user-interface events to handle a ribbon reset.
            UiEvents = ThisApplication.UserInterfaceManager.UserInterfaceEvents

        End Sub

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            UiEvents = Nothing
            ThisApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other 
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:this method is now obsolete, you should use the 
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub

#End Region

#Region "User interface definition"
        ' Sub where the user-interface creation is done.  This is called when
        ' the add-in loaded and also if the user interface is reset.
        Private Sub AddToUserInterface()
            ' Add DXF Update to Sheet Metal Tools tab
            Dim partRibbon As Ribbon = ThisApplication.UserInterfaceManager.Ribbons.Item("Part")
            Dim toolsTab As RibbonTab = partRibbon.RibbonTabs.Item("id_TabSheetMetal")
            Dim customPanel As RibbonPanel = toolsTab.RibbonPanels.Add("Add-Ins", "dxfUpdate", AddInClientID)
            customPanel.CommandControls.AddButton(DXFUpdate, True) ' True = Large icon

            ' Add Print Update to Drawing Place Views tab
            partRibbon = ThisApplication.UserInterfaceManager.Ribbons.Item("Drawing")
            toolsTab = partRibbon.RibbonTabs.Item("id_TabPlaceViews")
            customPanel = toolsTab.RibbonPanels.Add("Add-Ins", "printUpdate", AddInClientID)
            customPanel.CommandControls.AddButton(PrintUpdate, True) ' True = Large icon

            ' Add Print Update to Drawing Annotate tab
            Dim annotateTab As RibbonTab = partRibbon.RibbonTabs.Item("id_TabAnnotate")
            Dim annotatePanel As RibbonPanel = annotateTab.RibbonPanels.Add("Add-Ins", "printUpdateAnnotate", AddInClientID)
            annotatePanel.CommandControls.AddButton(PrintUpdate, True) ' True = Large icon
        End Sub

        Private Sub UiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles UiEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

        ' Sample handler for the button.
        Private Shared Sub DXFUpdate_OnExecute(Context As NameValueMap) Handles DXFUpdate.OnExecute
            Call Sub() runDxfUpdate()
            'Call Sub() userName()
        End Sub

        Private Shared Sub PrintUpdate_OnExecute(Context As NameValueMap) Handles PrintUpdate.OnExecute
            Call Sub() RunPrintUpdate()
        End Sub
#End Region

    End Class
End Namespace


Public Module Globals

#Region "Function to get the add-in client ID."
    ' This function uses reflection to get the GuidAttribute associated with the add-in.
    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(DoyleAddin.StandardAddInServer)
            Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
            Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
            guid = "{" + guidAttribute.Value.ToString() + "}"
        Catch
        End Try

        Return guid
    End Function
#End Region

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(ThisApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private ReadOnly _hwnd As IntPtr
    End Class
#End Region

End Module