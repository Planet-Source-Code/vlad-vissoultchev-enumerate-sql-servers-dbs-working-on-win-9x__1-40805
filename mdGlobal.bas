Attribute VB_Name = "mdGlobal"
Option Explicit
Private Const MODULE_NAME As String = "mdGlobal"

Private Sub ShowError(sFunc As String)
    Screen.MousePointer = vbDefault
    MsgBox "Error: 0x" & Hex(Err.Number) & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "Call stack:" & vbCrLf & MODULE_NAME & "." & sFunc & vbCrLf & Err.Source, vbCritical
End Sub

Sub Main()
    Const FUNC_NAME     As String = "Main"
    Const SEPARATOR     As String = ""
    Dim vSplit          As Variant
    Dim sSrv            As String
    Dim sDb             As String
    Dim sUser           As String
    Dim sPass           As String
    
    On Error GoTo EH
    '--- restore settings from registry
    vSplit = Split(GetSetting("EnumSqlDbs", "Common", "Settings", ""), SEPARATOR)
    If UBound(vSplit) >= 3 Then
        sSrv = vSplit(0)
        sDb = vSplit(1)
        sUser = vSplit(2)
        sPass = vSplit(3)
    End If
    '--- if confirmed -> store settings in registry
    If frmDbSettings.Init(sSrv, sDb, sUser, sPass) Then
        SaveSetting "EnumSqlDbs", "Common", "Settings", sSrv & SEPARATOR & sDb & SEPARATOR & sUser & SEPARATOR & sPass
    End If
    Exit Sub
EH:
    ShowError FUNC_NAME
End Sub
