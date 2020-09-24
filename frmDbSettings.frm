VERSION 5.00
Begin VB.Form frmDbSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Settings"
   ClientHeight    =   3108
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   5124
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   7.8
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDbSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3108
   ScaleWidth      =   5124
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   348
      Left            =   3864
      TabIndex        =   11
      Top             =   2604
      Width           =   1104
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   348
      Left            =   2688
      TabIndex        =   10
      Top             =   2604
      Width           =   1104
   End
   Begin VB.ComboBox cobDbAdo 
      Height          =   288
      Left            =   2268
      TabIndex        =   8
      Top             =   1596
      Width           =   2700
   End
   Begin VB.TextBox txtPass 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   2268
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1176
      Width           =   2700
   End
   Begin VB.TextBox txtUser 
      Height          =   288
      Left            =   2268
      TabIndex        =   5
      Top             =   756
      Width           =   2700
   End
   Begin VB.ComboBox cobDbOdbc 
      Height          =   288
      Left            =   2268
      TabIndex        =   1
      Top             =   2016
      Width           =   2700
   End
   Begin VB.ComboBox cobServer 
      Height          =   288
      Left            =   2268
      TabIndex        =   0
      Top             =   336
      Width           =   2700
   End
   Begin VB.Label labIntegrated 
      Caption         =   "(Integrated Security)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2268
      TabIndex        =   12
      Top             =   1176
      Width           =   2448
   End
   Begin VB.Label Label5 
      Caption         =   "Database (ADO):"
      Height          =   264
      Left            =   588
      TabIndex        =   9
      Top             =   1596
      Width           =   1692
   End
   Begin VB.Label Label4 
      Caption         =   "Database (ODBC):"
      Height          =   264
      Left            =   588
      TabIndex        =   7
      Top             =   2016
      Width           =   1692
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   264
      Left            =   588
      TabIndex        =   4
      Top             =   1176
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "User:"
      Height          =   264
      Left            =   588
      TabIndex        =   3
      Top             =   756
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Server:"
      Height          =   264
      Left            =   588
      TabIndex        =   2
      Top             =   336
      Width           =   1692
   End
End
Attribute VB_Name = "frmDbSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "frmDbSettings"

Private m_bEnumSrv          As Boolean
Private m_bEnumDbOdbc       As Boolean
Private m_bEnumDbAdo        As Boolean
Private m_bOk               As Boolean

Private Sub ShowError(sFunc As String)
    Screen.MousePointer = vbDefault
    MsgBox "Error: 0x" & Hex(Err.Number) & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "Call stack:" & vbCrLf & MODULE_NAME & "." & sFunc & vbCrLf & Err.Source, vbCritical
End Sub

Private Sub RaiseError(sFunc As String)
    Err.Raise Err.Number, MODULE_NAME & "." & sFunc & vbCrLf & Err.Source, Err.Description
End Sub

Public Function Init( _
            sSrv As String, _
            sDb As String, _
            sUser As String, _
            sPass As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    
    On Error GoTo EH
    cobServer = sSrv
    cobDbAdo = sDb
    txtUser = sUser
    txtUser_Change
    txtPass = sPass
    Show vbModal
    If m_bOk Then
        sSrv = cobServer
        sDb = cobDbAdo
        sUser = txtUser
        sPass = txtPass
        '--- success
        Init = True
    End If
    Unload Me
    Exit Function
EH:
    ShowError FUNC_NAME
    Resume NextLine
NextLine:
    On Error Resume Next
    Unload Me
End Function

Private Sub cmdCancel_Click()
    Visible = False
End Sub

Private Sub cmdOk_Click()
    m_bOk = True
    Visible = False
End Sub

Private Sub cobDbAdo_DropDown()
    Const FUNC_NAME     As String = "cobDbOdbc_DropDown"
    Dim vDb             As Variant
    Dim sText           As String
    
    On Error GoTo EH
    If Not m_bEnumDbAdo Then
        Screen.MousePointer = vbHourglass
        sText = cobDbOdbc.Text
        cobDbAdo.Clear
        For Each vDb In EnumSqlDbAdo(cobServer.Text, txtUser.Text, txtPass.Text)
            cobDbAdo.AddItem vDb
        Next
        cobDbOdbc.Text = sText
        m_bEnumDbAdo = True
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
EH:
    ShowError FUNC_NAME
End Sub

Private Sub cobDbOdbc_DropDown()
    Const FUNC_NAME     As String = "cobDbOdbc_DropDown"
    Dim vDb             As Variant
    Dim sText           As String
    
    On Error GoTo EH
    If Not m_bEnumDbOdbc Then
        Screen.MousePointer = vbHourglass
        sText = cobDbOdbc.Text
        cobDbOdbc.Clear
        For Each vDb In EnumSqlDbs(cobServer.Text, txtUser.Text, txtPass.Text)
            cobDbOdbc.AddItem vDb
        Next
        cobDbOdbc.Text = sText
        m_bEnumDbOdbc = True
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
EH:
    ShowError FUNC_NAME
End Sub

Private Sub cobServer_Change()
    m_bEnumDbOdbc = False
    m_bEnumDbAdo = False
End Sub

Private Sub cobServer_Click()
    m_bEnumDbOdbc = False
    m_bEnumDbAdo = False
End Sub

Private Sub cobServer_DropDown()
    Const FUNC_NAME     As String = "cobServer_DropDown"
    Dim vSrv            As Variant
    
    On Error GoTo EH
    If Not m_bEnumSrv Then
        Screen.MousePointer = vbHourglass
        For Each vSrv In EnumSqlServers
            cobServer.AddItem vSrv
        Next
        m_bEnumSrv = True
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
EH:
    ShowError FUNC_NAME
End Sub

Private Sub txtPass_Change()
    m_bEnumDbOdbc = False
    m_bEnumDbAdo = False
End Sub

Private Sub txtUser_Change()
    m_bEnumDbOdbc = False
    m_bEnumDbAdo = False
    txtPass.Visible = Len(txtUser) > 0
End Sub
