VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9135
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10800
   _ExtentX        =   19050
   _ExtentY        =   16113
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "PSC Explorer"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed As Boolean
Public VBInstance As VBIDE.VBE
Public gWinWindow As VBIDE.Window
Dim mcbMenuCommandBar As Office.CommandBarControl
Dim mcbToolCommandBar As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents TBHandler As CommandBarEvents
Attribute TBHandler.VB_VarHelpID = -1
Private docUI As udMain


'The Dockable Connect.dsr modified from Alfred YomTov's Code Finder
'http://www.pscode.com/vb/default.asp?lngCId=37606&lngWId=1


Sub Run()
  Const WindowGUID = "PSC_EXPLORER"
  On Error Resume Next

  If gWinWindow Is Nothing Then
    Set gWinWindow = VBInstance.Windows.CreateToolWindow(VBInstance.Addins("PSCExplorer.Connect"), "PSCExplorer.udMain", _
        "PSC Explorer v" & App.Major & "." & App.Minor & "." & App.Revision, WindowGUID, docUI)

  End If

  Set docUI.Connect = Me
  FormDisplayed = True
  gWinWindow.Visible = True


End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
  On Error GoTo error_handler

  'save the vb instance
  Set VBInstance = Application

  'this is a good place to set a breakpoint and
  'test various addin objects, properties and methods
  'Debug.Print VBInstance.FullName

  If ConnectMode = ext_cm_External Then
    'Used by the wizard toolbar to start this wizard
    Call Run
  Else
    Set mcbMenuCommandBar = AddToAddInCommandBar("PSC Explorer")
    Set mcbToolCommandBar = AddToToolBarCommandBar("PSC Explorer")

    'sink the event
    Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    Set Me.TBHandler = VBInstance.Events.CommandBarEvents(mcbToolCommandBar)
    
  End If

  If ConnectMode = ext_cm_AfterStartup Then
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
      'set this to display the form on connect
      Call Run
    End If
  End If

  Exit Sub

error_handler:

  MsgBox Err.Description

End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
  On Error Resume Next

  'delete the command bar entry
  mcbMenuCommandBar.Delete
  mcbToolCommandBar.Delete
  'shut down the Add-In
  If FormDisplayed Then
    SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
    FormDisplayed = False
  Else
    SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
  End If

  Set gWinWindow = Nothing


End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
  If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
    'set this to display the form on connect
    Call Run
  End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Run
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
  Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
  Dim cbMenu As Object

  On Error GoTo AddToAddInCommandBarErr

  'see if we can find the Add-Ins menu
  Set cbMenu = VBInstance.CommandBars("Add-Ins")
  If cbMenu Is Nothing Then
    'not available so we fail
    Exit Function
  End If

  'add it to the command bar
  Set cbMenuCommandBar = cbMenu.Controls.Add(1)
  'set the caption
  cbMenuCommandBar.Caption = sCaption
  ' Set the picture of the button by copying resource
  ' and then pasting it on to it.
  Clipboard.Clear
  Clipboard.SetData LoadResPicture("MENUPIC2", vbResBitmap), vbCFBitmap
  cbMenuCommandBar.PasteFace
  Set AddToAddInCommandBar = cbMenuCommandBar

  Exit Function

AddToAddInCommandBarErr:

End Function

Function AddToToolBarCommandBar(sCaption As String) As Office.CommandBarControl
  Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
  Dim cbMenu As Object

  On Error GoTo AddToToolBarCommandBarErr

  'see if we can find the Add-Ins menu
  Set cbMenu = VBInstance.CommandBars("Standard")
  If cbMenu Is Nothing Then
    'not available so we fail
    Exit Function
  End If

  'add it to the command bar
  Set cbMenuCommandBar = cbMenu.Controls.Add(1, , , 20)

  'set the caption
  cbMenuCommandBar.Caption = sCaption
  cbMenuCommandBar.ToolTipText = sCaption
  ' Set the picture of the button by copying resource
  ' and then pasting it on to it.
  Clipboard.Clear
  Clipboard.SetData LoadResPicture("MENUPIC2", vbResBitmap), vbCFBitmap
  cbMenuCommandBar.PasteFace
  Set AddToToolBarCommandBar = cbMenuCommandBar

  Exit Function

AddToToolBarCommandBarErr:

End Function

Private Sub TBHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Run
End Sub
