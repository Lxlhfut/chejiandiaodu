VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmorder1 
   Caption         =   "����ά������"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "����ά��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   14775
      Begin VB.TextBox cmblocomotivetype 
         Height          =   375
         Left            =   2760
         TabIndex        =   24
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox txtworkcode 
         Height          =   330
         Left            =   2760
         TabIndex        =   23
         Top             =   960
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   10680
         TabIndex        =   22
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57540611
         CurrentDate     =   37056
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   10680
         TabIndex        =   21
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57540611
         CurrentDate     =   37056
      End
      Begin VB.TextBox Txtamount 
         Height          =   315
         Left            =   6600
         TabIndex        =   18
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Txtproductname 
         Height          =   315
         Left            =   6600
         TabIndex        =   16
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtorder 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtdrawingnumber 
         Height          =   315
         Left            =   6600
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtnote 
         Height          =   330
         Left            =   10680
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   210
         Left            =   9600
         TabIndex        =   20
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ������"
         Height          =   210
         Left            =   9600
         TabIndex        =   19
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʒ����"
         Height          =   210
         Left            =   5640
         TabIndex        =   17
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʒ����"
         Height          =   210
         Left            =   5640
         TabIndex        =   15
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   210
         Left            =   2100
         TabIndex        =   14
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   210
         Left            =   2100
         TabIndex        =   13
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ͺ�"
         Height          =   210
         Left            =   1860
         TabIndex        =   12
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʒͼ��"
         Height          =   210
         Left            =   5640
         TabIndex        =   11
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��   ע"
         Height          =   210
         Left            =   9720
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "���"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   2160
      Width           =   1635
   End
   Begin VB.CommandButton cmd_select 
      Caption         =   "��ѯ"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2160
      Width           =   1635
   End
   Begin VB.CommandButton cmd_renew 
      Caption         =   "�޸�"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   2160
      Width           =   1635
   End
   Begin VB.CommandButton cmd_delete 
      Caption         =   "ɾ��"
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   2160
      Width           =   1635
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   11760
      TabIndex        =   0
      Top             =   2160
      Width           =   1635
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   12515
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frmorder1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim bm

Private Sub cmd_add_Click()
    cmd_renew.Enabled = False
    cmd_delete.Enabled = False
    
    Dim sql As String
    Dim rst As New ADODB.Recordset
    
        If Trim(txtorder) = "" Then
            MsgBox "�����Ų���Ϊ��ֵ", vbOKOnly + vbInformation, "��ʾ"
            cmd_renew.Enabled = True
            cmd_delete.Enabled = True
            Exit Sub
        End If
        If Trim(txtworkcode) = "" Then
            MsgBox "�����Ų���Ϊ��ֵ", vbOKOnly + vbInformation, "��ʾ"
            cmd_renew.Enabled = True
            cmd_delete.Enabled = True
            Exit Sub
        End If
        If Trim(cmblocomotivetype) = "" Then
            MsgBox "�����ͺŲ���Ϊ��ֵ", vbOKOnly + vbInformation, "��ʾ"
            cmd_renew.Enabled = True
            cmd_delete.Enabled = True
            Exit Sub
        End If
        If Trim(txtdrawingnumber) = "" Then
            MsgBox "��Ʒͼ�Ų���Ϊ��ֵ", vbOKOnly + vbInformation, "��ʾ"
            cmd_renew.Enabled = True
            cmd_delete.Enabled = True
            Exit Sub
        End If
        
        Set rst = Nothing
        sql = "select * from t_suborder where ordercode='" & Trim(txtorder) & "'" & _
                " and drawingnumber='" & Trim(txtdrawingnumber) & "'"
        rst.CursorLocation = adUseClient
        rst.Open sql, conn, adOpenKeyset, adLockPessimistic
    If rst.RecordCount <> 0 Then
            MsgBox "�ζ����Ѿ����룬�����²���..."
            cmd_renew.Enabled = True
            cmd_delete.Enabled = True
            Exit Sub
    Else
        rs.AddNew
        rs("ordercode") = Trim(txtorder)
        rs("workcode") = Trim(txtworkcode)
        rs("locomotivetype") = Trim(cmblocomotivetype)
        rs("drawingnumber") = Trim(txtdrawingnumber)
        rs("productname") = Trim(Txtproductname)
        rs("amount") = Trim(Txtamount)
        rs("acceptdate") = DTPicker1.Value
        rs("senddate") = DTPicker2.Value
        rs("added") = "��"
        rs("note") = Trim(txtnote)
            
            Dim yn As Integer
            yn = MsgBox("ȷ��������ȷ��?", vbYesNo + vbQuestion)
            
            If yn = vbYes Then
               rs.Update
            Else
               rs.CancelUpdate
            End If
     End If
            cmd_renew.Enabled = True
            cmd_delete.Enabled = True
    Set rst = Nothing
    Exit Sub
errhealer:
    MsgBox err.Description & err.number, vbOKOnly + vbInformation, "������ʾ"
    rs.CancelUpdate
    err.Clear
End Sub
Private Sub cmd_delete_Click()
    Dim yn As Integer
    yn = MsgBox("�����Ҫɾ��������¼��", vbYesNo + vbQuestion)
    If yn = vbYes Then
        rs.delete
        DoEvents
        Call init
        If rs.RecordCount = 0 Then
            cmd_renew.Enabled = False
            cmd_delete.Enabled = False
        Else
             DoEvents
        End If
     End If
            cmd_renew.Enabled = True
            cmd_delete.Enabled = True
End Sub

Private Sub cmd_exit_Click()
    Unload Me
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
End Sub

Private Sub cmd_renew_Click()
    On Error GoTo errhealer
    Dim yn As Integer
    
    If Trim(txtorder) = "" Then
            MsgBox "�����Ų���Ϊ��ֵ", vbOKOnly + vbInformation, "��ʾ"
            Exit Sub
        End If
             
        If Trim(txtworkcode) = "" Then
            MsgBox "�����Ų���Ϊ��ֵ", vbOKOnly + vbInformation, "��ʾ"
            Exit Sub
        End If

        If Trim(cmblocomotivetype) = "" Then
            MsgBox "�����ͺŲ���Ϊ��ֵ", vbOKOnly + vbInformation, "��ʾ"
            Exit Sub
        End If

        If Trim(txtdrawingnumber) = "" Then
            MsgBox "��Ʒͼ�Ų���Ϊ��ֵ", vbOKOnly + vbInformation, "��ʾ"
            Exit Sub
        End If
       ' rs.AddNew
        rs("ordercode") = Trim(txtorder)
        rs("workcode") = Trim(txtworkcode)
        rs("locomotivetype") = Trim(cmblocomotivetype)
        rs("drawingnumber") = Trim(txtdrawingnumber)
        rs("productname") = Trim(Txtproductname)
        rs("amount") = Trim(Txtamount)
        rs("acceptdate") = DTPicker1.Value
        rs("senddate") = DTPicker2.Value
        rs("added") = "��"
        rs("note") = Trim(txtnote)
    yn = MsgBox("ȷ��������ȷ��?", vbYesNo + vbQuestion)
    If yn = vbYes Then
        rs.Update
    Else
      DoEvents
    End If
    Exit Sub
errhealer:
    MsgBox err.Description, vbOKOnly + vbInformation, "������ʾ"
    rs.CancelUpdate
    err.Clear
End Sub

Private Sub cmd_select_Click()
    Dim sql As String
    Dim pos As Integer
    If rs.State <> adStateClosed Then
        Set rs = Nothing
    End If
            sql = "select * from t_suborder  "
            
            If Trim(txtorder) <> "" Then
               sql = sql & " where  ordercode= '" & Trim(txtorder) & "'"
            End If
            
            If Trim(txtdrawingnumber) <> "" Then
                pos = InStr(sql, "where")
                If pos <> 0 Then
                   sql = sql & " and drawingnumber='" & Trim(txtdrawingnumber) & "'"
                Else
                   sql = sql & " where drawingnumber='" & Trim(txtdrawingnumber) & "'"
                End If
            End If
             rs.CursorLocation = adUseClient
             rs.Open sql, conn, adOpenKeyset, adLockPessimistic
             
             If rs.RecordCount = 0 Then
                MsgBox "�Բ����Ҳ�����Ӧ�ļ�¼", vbOKOnly + vbInformation
             End If
             
             Set DataGrid1.DataSource = rs
             DataGrid1.Refresh
             Call setgrid(DataGrid1)
End Sub

Private Sub DataGrid1_Click()
    If rs.RecordCount <> 0 Then
        rs.Bookmark = DataGrid1.Bookmark
        txtorder = rs("ordercode")
        txtworkcode = rs("workcode")
        cmblocomotivetype = rs("locomotivetype")
        txtdrawingnumber = rs("drawingnumber")
        Txtproductname = rs("productname")
        Txtamount = rs("amount")
        'Cmbfou = rs("added")
        If Not IsNull(rs("note")) Then txtnote = rs("note")
    End If
End Sub

Private Sub Form_Load()
    Dim sql
    Dim rst As New ADODB.Recordset
    Dim i As Integer
    Call init
    conn.Open "dsn=dbw", "sa"
    
    Set rs = Nothing
    sql = "select * from t_suborder "
    rs.CursorLocation = adUseClient
    rs.Open sql, conn, adOpenKeyset, adLockPessimistic
    Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
    Call setgrid(DataGrid1)
    DTPicker1.Value = Now()
    DTPicker2.Value = Now()
    
    Set rst = Nothing
    sql = "select * from t_subworkcode "
    rst.CursorLocation = adUseClient
    rst.Open sql, conn, adOpenKeyset, adLockOptimistic
    If rst.RecordCount <> 0 Then
            rst.MoveFirst
            For i = 0 To rst.RecordCount - 1
                txtworkcode.AddItem rst("workcode")
                rst.MoveNext
            Next
     End If
     Set rst = Nothing
End Sub
Private Sub init()
    Dim ctl As Control
    For Each ctl In Me.Controls
         If TypeOf ctl Is TextBox Then
            ctl = ""
        End If
    Next ctl
End Sub
Private Sub setgrid(dg As DataGrid)
      Dim i As Integer
      Dim pwidth As Integer
      dg.Columns.Item(0).Caption = "������"
      dg.Columns.Item(1).Caption = "������"
      dg.Columns.Item(2).Caption = "�����ͺ�"
      dg.Columns.Item(3).Caption = "��Ʒͼ��"
      dg.Columns.Item(4).Caption = "��Ʒ����"
      dg.Columns.Item(5).Caption = "��Ʒ����"
      dg.Columns.Item(6).Caption = "Ԥ������"
      dg.Columns.Item(7).Caption = "��������"
      dg.Columns.Item(8).Caption = "����ƻ���"
      dg.Columns.Item(9).Caption = "��ע"
      pwidth = Fix((dg.Width - 600) / 10)
      For i = 0 To 9
         dg.Columns.Item(i).Width = pwidth
      Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call cmd_exit_Click
End Sub

Private Sub txtdrawingnumber_LostFocus()
    Dim rst As New ADODB.Recordset
    Dim sql As String
    Set rst = Nothing
    sql = "select * from t_spbillofmaterial where pardrawingnumber='" & Trim(txtdrawingnumber) & "'"
    rst.CursorLocation = adUseClient
    rst.Open sql, conn, adOpenKeyset, adLockOptimistic
    If rst.RecordCount <> 0 Then
        Txtproductname = Trim(rst("partname"))
    Else
       MsgBox "û�����ͼ�����Ӧ�Ĳ�Ʒ", vbExclamation + vbInformation
       Exit Sub
    End If
    Set rst = Nothing
End Sub

Private Sub txtworkcode_LostFocus()
    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    Set rst = Nothing
    sql = "select * from t_subworkcode where workcode='" & Trim(txtworkcode.Text) & "'"
    rst.CursorLocation = adUseClient
    rst.Open sql, conn, adOpenKeyset, adLockOptimistic
    If rst.RecordCount <> 0 Then
            cmblocomotivetype.Text = CStr(rst("productname"))
     End If
     Set rst = Nothing
     End Sub
