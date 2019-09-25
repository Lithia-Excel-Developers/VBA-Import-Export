Attribute VB_Name = "modMenu"
Option Explicit

Private MnuEvt      As clsVBECmdHandler
Private EvtHandlers As New Collection

Private Const STR_DEFAULT_VCSMENUCAPTION As String = "Version &Control"
Private Const STR_DEFAULT_TESTMENUCAPTION As String = "Te&sts"

Public Sub auto_open()

    CreateVBEVCSMenu
    CreateVBETestMenu

End Sub


Public Sub auto_close()
    
    RemoveVBEMenu STR_DEFAULT_VCSMENUCAPTION
    RemoveVBEMenu STR_DEFAULT_TESTMENUCAPTION

End Sub


Private Sub CreateVBEVCSMenu()

    Dim objMenu     As CommandBarPopup
    Dim objMenuItem As Object

    Set objMenu = Application.VBE.CommandBars(1).Controls.Add(Type:=msoControlPopup)
    With objMenu
        objMenu.Caption = STR_DEFAULT_VCSMENUCAPTION

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "MakeConfigFile"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Update Config File"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "Import"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Import"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "Export"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Export"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "Save"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Save"

    End With

End Sub

Private Sub CreateVBETestMenu()

    Dim objMenu     As CommandBarPopup
    Dim objMenuItem As Object

    Set objMenu = Application.VBE.CommandBars(1).Controls.Add(Type:=msoControlPopup)
    With objMenu
        objMenu.Caption = STR_DEFAULT_TESTMENUCAPTION

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "RunTests"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Run All Tests"
        
        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "InitTests"
        MenuEvents objMenuItem
        objMenuItem.Caption = "Add Tests to &Config"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "ImportTests"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Import"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "ExportTests"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Export"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "Save"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Save"

    End With
       
End Sub


Private Sub MenuEvents(ByVal objMenuItem As Object)

    Set MnuEvt = New clsVBECmdHandler
    Set MnuEvt.EvtHandler = Application.VBE.Events.CommandBarEvents(objMenuItem)
    EvtHandlers.Add MnuEvt

End Sub


Private Sub RemoveVBEMenu(ByVal MenuName As String)

    On Error Resume Next

    Application.VBE.CommandBars(1).Controls(Replace(MenuName, "&", "")).Delete

    '// Clear the EvtHandlers collection if there is anything in it
    While EvtHandlers.Count > 0
        EvtHandlers.Remove 1
    Wend

    Set EvtHandlers = Nothing
    Set MnuEvt = Nothing

    Application.CommandBars("Worksheet Menu Bar").Controls(Replace(MenuName, "&", "")).Delete
    On Error GoTo 0

End Sub

'// RibUI callbacks
Public Sub btnMakeConfig_onAction(control As IRibbonControl)
    MakeConfigFile
End Sub
Public Sub btnExport_onAction(control As IRibbonControl)
    Export
End Sub
Public Sub btnSave_onAction(control As IRibbonControl)
    Save
End Sub
Public Sub btnImport_onAction(control As IRibbonControl)
    Import
End Sub
Public Sub btnRunTests_onAction(control As IRibbonControl)
    RunTests
End Sub

