VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Tag             =   "��¼"
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   360
      Left            =   2100
      TabIndex        =   4
      Tag             =   "ȡ��"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   360
      Left            =   495
      TabIndex        =   3
      Tag             =   "ȷ��"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   270
      Left            =   1305
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "����(&P):"
      Height          =   248
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Tag             =   "����(&P):"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "�û���(&U):"
      Height          =   248
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Tag             =   "�û���(&U):"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Dim mconn As New ADODB.Connection
Public OK As Boolean
Private Sub Form_Load()
    Dim sBuffer As String
    Dim lSize As Long
'    mconn.Open "DSN=dlrwdb;uid=sa" ';uid=scl;uid=scl"
     mconn.Open "DSN=dbw;uid=sa" ';uid=scl"
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    'Call GetUserName(sBuffer, lSize)
    sBuffer = "scl"
    txtPassword.Text = "scl"
    If lSize > 0 Then
        txtUserName.Text = Left$(sBuffer, lSize)
    Else
        txtUserName.Text = vbNullString
    End If
End Sub



Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    'ToDo: �������������Ƿ���ȷ
    '�����ȷ����
Dim sql As String, rs As New ADODB.Recordset
sql = "select * from passwd where opnumber='" & Trim$(txtUserName) & "'"
Set rs = Nothing
rs.CursorLocation = adUseClient
rs.Open sql, mconn, adOpenKeyset, adLockPessimistic
If rs.RecordCount = 0 Then
    MsgBox "no user", vbOKOnly
    rs.Close
    Exit Sub
End If
    'ToDo: �������������Ƿ���ȷ
    '�����ȷ����
    If txtPassword.Text = Trim$(rs("pass")) Then
        OK = True
        CurrentUser = txtUserName.Text
        rs.Close
        Me.Hide
'        Call init_env
    Else
        MsgBox "�����������һ�Σ�", , "��¼"
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If

End Sub

Sub init_env()
Dim sql As String, rs As New ADODB.Recordset
sql = "select * from t_ctrl"
rs.CursorLocation = adUseClient
rs.Open sql, mconn, adOpenKeyset, adLockPessimistic
rs.MoveFirst
Period = rs("period")
rs.Close
End Sub

