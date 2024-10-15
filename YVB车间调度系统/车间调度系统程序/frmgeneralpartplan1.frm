VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmgeneralpartplan1 
   Caption         =   "计划拆解及排序"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleMode       =   0  'User
   ScaleWidth      =   10000
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   18865
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "订单计划拆解"
      TabPicture(0)   =   "frmgeneralpartplan1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DataGrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DataGrid4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "按调度规则排序"
      TabPicture(1)   =   "frmgeneralpartplan1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid3"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   5895
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "与工艺对照的零件计划表"
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   6720
         Width           =   14295
         Begin VB.CommandButton Comexit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "退出"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton comshchjh 
            BackColor       =   &H00C0C0C0&
            Caption         =   "生成零件计划"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6120
            TabIndex        =   8
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmbfind 
            BackColor       =   &H00C0C0C0&
            Caption         =   "查询未加工计划"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   7
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Height          =   1935
         Left            =   -74760
         TabIndex        =   1
         Top             =   6720
         Width           =   13575
         Begin VB.CommandButton Cmdfresh 
            Caption         =   "刷 新"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10560
            TabIndex        =   17
            Top             =   840
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Height          =   1215
            Left            =   360
            TabIndex        =   12
            Top             =   480
            Width           =   9495
            Begin VB.CommandButton Cmddel 
               BackColor       =   &H00C0C0C0&
               Caption         =   "删除"
               Height          =   375
               Left            =   7080
               TabIndex        =   16
               Top             =   600
               Width           =   1335
            End
            Begin VB.ComboBox cmbddgz 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmgeneralpartplan1.frx":0038
               Left            =   840
               List            =   "frmgeneralpartplan1.frx":003A
               TabIndex        =   14
               Top             =   720
               Width           =   1935
            End
            Begin VB.ListBox List1 
               Height          =   420
               ItemData        =   "frmgeneralpartplan1.frx":003C
               Left            =   3960
               List            =   "frmgeneralpartplan1.frx":003E
               TabIndex        =   13
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "调度规则"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   1320
               TabIndex        =   15
               Top             =   360
               Width           =   840
            End
         End
         Begin VB.CommandButton Comddgz 
            BackColor       =   &H00C0C0C0&
            Caption         =   "调度排序"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10560
            TabIndex        =   3
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Comexit2 
            BackColor       =   &H8000000B&
            Caption         =   "退出"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10560
            TabIndex        =   2
            Top             =   1320
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   5775
         Left            =   -74520
         TabIndex        =   4
         Top             =   480
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   10186
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   7800
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "未加工任务"
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   6135
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   10821
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "本月零件计划"
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
End
Attribute VB_Name = "frmgeneralpartplan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim rsr As New ADODB.Recordset
Dim rss As New ADODB.Recordset '向设备分配任务
Dim rs2 As New ADODB.Recordset  '用来记录零件计划
Dim total As Integer
Dim myrs As New ADODB.Recordset '用来记录生成的零部件计划
Dim mark
Dim bool As Boolean
Dim C() As Integer
Dim conn As New ADODB.Connection


'根据零部件的工作号、机车型号,图号,计划个数、台分数量、生产类型查找以他为父件的零部件
Public Sub ff(workcodes As String, ordercodes As String, locomotivetypes As String, prodrawings As String, progroupamounts As Integer, groupamounts As Integer)
   Dim sql As String
   Dim bmrs As New ADODB.Recordset  '用来记录生成的零部件计划
   Dim find As Boolean '表示本计划再零部件计划表中是否存在
   Dim find1 As Boolean '表示本记录的借用关系
   Dim sch As New ADODB.Recordset
   Dim ddrs As New ADODB.Recordset    '他为查询大纲计划是否在调度模式计划中的记录集
   Dim orderrs As New ADODB.Recordset
   sql = "select * from t_spbillofmaterial" & _
         " where locomotivetype='" & Trim$(locomotivetypes) & "'" & _
        "  and prodrawingnumber='" & Trim$(prodrawings) & "'"  '  and   productiontype in ('委外加工','自制','组装')"
  bmrs.ActiveConnection = "dsn=dbw;uid=sa"
  bmrs.CursorLocation = adUseClient
  bmrs.CursorType = adOpenDynamic
  bmrs.LockType = adLockOptimistic
  bmrs.Source = sql
  bmrs.Open
 
  If bmrs.EOF And bmrs.BOF Then  '若没有此部件的产品明细,则退出
         bmrs.Close
        Set bmrs = Nothing
       ' mark = mark + 1
        Exit Sub
  End If '否则转换成另不见计划

    On Error Resume Next
    bmrs.MoveFirst
    Do Until bmrs.EOF '首先判断本条记录是否为调度模式计划若是则去掉

                 If bmrs("pargroupamount") <> 0 Then
                     If rs2.RecordCount <> 0 Then '首先判断零部件计划表中是否有本计划,若有则只增加其数量,否则增加一条新记录.
                           find = False
                              sql = "select * from t_spgeneralpartplan" & _
                                    "  where  ordercode='" & Trim$(ordercodes) & "'and workcode='" & Trim$(workcodes) & "' and  drawingnumber='" & Trim$(bmrs("pardrawingnumber")) & "'"
                         
                              If sch.State = adStateOpen Then sch.Close
                              sch.ActiveConnection = "dsn=dbw;uid=sa"
                              sch.CursorLocation = adUseClient
                              sch.CursorType = adOpenKeyset
                              sch.LockType = adLockOptimistic
                              sch.Source = sql
                              sch.Open
                              If sch.RecordCount <> 0 Then find = True
                      End If
                     
                    If find Then
                          sch("planquantity") = sch("planquantity") + progroupamounts * bmrs("pargroupamount") / groupamounts '父件计划个数*子件台分数量/父件台分数量
                          sch.Update
                     Else
                         Set orderrs = Nothing
                         orderrs.ActiveConnection = "dsn=dbw;uid=sa"
                         orderrs.CursorLocation = adUseClient
                         orderrs.CursorType = adOpenKeyset
                         orderrs.LockType = adLockOptimistic
                         orderrs.Source = "select * from t_suborder where  ordercode='" & Trim$(ordercodes) & "'"
                         orderrs.Open
                        If orderrs.RecordCount <> 0 Then orderrs.MoveFirst
                         rs2.AddNew
                         rs2("workcode") = Trim$(workcodes)
                         rs2("locomotivetype") = CStr(Trim$(bmrs("locomotivetype")))
                         rs2("ordercode") = Trim$(ordercodes)
                         rs2("prodrawingnumber") = CStr(Trim$(bmrs("prodrawingnumber")))
                         rs2("drawingnumber") = CStr(Trim$(bmrs("pardrawingnumber")))
                         rs2("planquantity") = progroupamounts * bmrs("pargroupamount") / groupamounts '父件计划个数*子件台分数量/父件台分数量
                         rs2("pargroupamount") = Int(bmrs("pargroupamount"))
                         rs2("acceptdate") = CDate(orderrs("acceptdate"))
                         rs2("senddate") = CDate(orderrs("senddate"))
                         rs2.Update
                         
                     End If
                            myrs.AddNew
                               myrs("workcode") = Trim$(workcodes)
                               myrs("ordercode") = Trim$(ordercodes)
                               myrs("locomotivetype") = CStr(Trim$(bmrs("locomotivetype")))
                               myrs("drawingnumber") = CStr(Trim$(bmrs("pardrawingnumber")))
                               myrs("pargroupamount") = Int(bmrs("pargroupamount"))
                               myrs("planquantity") = progroupamounts * Int(bmrs("pargroupamount")) / groupamounts
                               myrs.Update
                
           End If
       bmrs.MoveNext
    Loop
   bmrs.Close
   Set bmrs = Nothing
End Sub
'一产品的工作号、机车型号、图号、计划台分为参数进行拆零函数
Public Sub Main1(workcodes As String, ordercodes As String, locomotivetypes As String, pardrawings As String, pargroupamounts As Integer)   'groupamounts As Integer
   Dim sql As String
   Dim rs1 As New ADODB.Recordset '用来记录生成的零部件计划
   Dim find As Boolean '表示本计划再零部件计划表中是否存在
   Dim find1 As Boolean '表示本记录的借用关系
   Dim sch As New ADODB.Recordset
   Dim orderrs As New ADODB.Recordset
     '本条计划拆零按正常情况拆零
            sql = "select * from t_spbillofmaterial where locomotivetype='" & Trim$(locomotivetypes) & "'  and  " & _
                     " pardrawingnumber='" & Trim$(pardrawings) & "'" '& " and  productiontype in ('委外加工','自制','组装') "
                  rs1.ActiveConnection = "dsn=dbw;uid=sa"
                  rs1.CursorLocation = adUseClient
                  rs1.CursorType = adOpenKeyset
                  rs1.LockType = adLockOptimistic
                  rs1.Source = sql
                  rs1.Open
                  bool = False
                  If rs1.RecordCount = 0 Then '若没有此部件的产品明细,则退出
                     MsgBox "没有此部件的产品明细", vbExclamation, "提示"
                    bool = True
                    rs1.Close
                    Set rs1 = Nothing
                    Exit Sub
                  End If
                  
                   On Error Resume Next
                    rs1.MoveFirst
                Do Until rs1.EOF
                   If rs1("pargroupamount") <> 0 Then
                           If rs2.RecordCount <> 0 Then '首先判断零部件计划表中是否有本计划,若有则只增加其数量,否则增加一条新记录.
                                    find = False
                                     sql = "select * from t_spgeneralpartplan " & _
                                           "  where  ordercode='" & Trim$(ordercodes) & "' and  workcode='" & Trim$(workcodes) & "' and  drawingnumber='" & Trim$(rs1("pardrawingnumber")) & "'"
                            
                                     If sch.State = adStateOpen Then sch.Close
                                     sch.ActiveConnection = "dsn=dbw;uid=sa"
                                     sch.CursorLocation = adUseClient
                                     sch.CursorType = adOpenKeyset
                                     sch.LockType = adLockOptimistic
                                     sch.Source = sql
                                     sch.Open
                                     If sch.RecordCount <> 0 Then find = True
                           End If
                       
                      If find Then '若发现零部件计划里本计划已存在则增加新计划计划
                            sch("planquantity") = sch("planquantity") + pargroupamounts * rs1("pargroupamount")
                            sch.Update
                      Else '否则增加新计划
                      Set orderrs = Nothing
                      orderrs.ActiveConnection = "dsn=dbw;uid=sa"
                      orderrs.CursorLocation = adUseClient
                      orderrs.CursorType = adOpenKeyset
                      orderrs.LockType = adLockOptimistic
                      orderrs.Source = "select * from t_suborder where  ordercode='" & Trim$(ordercodes) & "'"
                      orderrs.Open
                      
                            rs2.AddNew
                            rs2("workcode") = Trim$(workcodes)
                            rs2("ordercode") = Trim$(ordercodes)
                            rs2("locomotivetype") = CStr(Trim$(rs1("locomotivetype")))
                            rs2("prodrawingnumber") = CStr(Trim$(rs1("prodrawingnumber")))
                            rs2("drawingnumber") = CStr(Trim$(rs1("pardrawingnumber")))
                            rs2("pargroupamount") = Int(rs1("pargroupamount"))
                            rs2("planquantity") = pargroupamounts * Int(rs1("pargroupamount"))
                            rs2("acceptdate") = CDate(orderrs("acceptdate"))
                            rs2("senddate") = CDate(orderrs("senddate"))
                            rs2.Update
                      End If  '将计划加入临时表中
                      
                               myrs.AddNew
                               myrs("workcode") = Trim$(workcodes)
                               myrs("ordercode") = Trim$(ordercodes)
                               myrs("locomotivetype") = CStr(Trim$(rs1("locomotivetype")))
                               myrs("drawingnumber") = CStr(Trim$(rs1("pardrawingnumber")))
                               myrs("pargroupamount") = Int(rs1("pargroupamount"))
                               myrs("planquantity") = pargroupamounts * Int(rs1("pargroupamount"))
                                myrs.Update
                    End If
                              rs1.MoveNext
                Loop
     rs1.Close
  Set rs1 = Nothing
  End Sub


Private Sub cmbddgz_LostFocus()
List1.AddItem (cmbddgz.Text)

End Sub

Private Sub cmbfind_Click()
   Dim sql As String
   sql = "select * from t_suborder where added='否'"
   Set rs = Nothing
   rs.ActiveConnection = "dsn=dbw;uid=sa"
   rs.CursorLocation = adUseClient
   rs.CursorType = adOpenKeyset
   rs.LockType = adLockOptimistic
   rs.Source = sql
   rs.Open
   Set DataGrid1.DataSource = rs
   Call initial(DataGrid1, "定单表")
   Call first(DataGrid1)
 If rs.RecordCount <> 0 Then
    comshchjh.Enabled = True
 End If
End Sub

Private Sub Cmddel_Click()
List1.RemoveItem (List1.ListIndex)
End Sub


Private Sub cmdFresh_Click()
    Dim findrs As New ADODB.Recordset
  Set findrs = Nothing
  findrs.ActiveConnection = "dsn=dbw;uid=sa"
  findrs.CursorLocation = adUseClient
  findrs.CursorType = adOpenDynamic
  findrs.LockType = adLockOptimistic
  findrs.Source = "delete from t_submachineload"
  findrs.Open
  Set findrs = Nothing
  findrs.ActiveConnection = "dsn=dbw;uid=sa"
  findrs.CursorLocation = adUseClient
  findrs.CursorType = adOpenDynamic
  findrs.LockType = adLockOptimistic
  findrs.Source = "delete from t_subdaytaskplan"
  findrs.Open
  Set findrs = Nothing
  findrs.ActiveConnection = "dsn=dbw;uid=sa"
  findrs.CursorLocation = adUseClient
  findrs.CursorType = adOpenDynamic
  findrs.LockType = adLockOptimistic
  findrs.Source = "delete from t_spgeneralpartplan"
  findrs.Open
  Set findrs = Nothing
  findrs.ActiveConnection = "dsn=dbw;uid=sa"
  findrs.CursorLocation = adUseClient
  findrs.CursorType = adOpenDynamic
  findrs.LockType = adLockOptimistic
  findrs.Source = "delete from t_myplantask"
  findrs.Open
  Set findrs = Nothing
  findrs.ActiveConnection = "dsn=dbw;uid=sa"
  findrs.CursorLocation = adUseClient
  findrs.CursorType = adOpenDynamic
  findrs.LockType = adLockOptimistic
  findrs.Source = "update t_suborder  set  added='否'"
  findrs.Open
End Sub

Private Sub Comddgz_Click()
Dim rs1 As New ADODB.Recordset  '用rs来对零件计划标排序并显示之
    'machine()代表机器编号，machine2()代表机器效率machine1()代表机器负荷
    Dim machine() As String, machine1() As Single, machine2() As Single
    Dim timeoccupy  '记录每台设备的时间占用
    Dim rs2 As New ADODB.Recordset '找出相应的产品对应的加工设备号
    Dim rs3 As New ADODB.Recordset '打开打开日生产计划表
    Dim rs5 As New ADODB.Recordset '用来给r表里添加数据
    Dim m As Integer, i As Integer, j As Integer
    Dim quota As Single
    Dim sql As String
    Dim s As String, s1 As String
    s = ""
    s1 = ""
    Dim rst As New ADODB.Recordset
    
    For j = 0 To List1.ListCount - 1
            Set rst = Nothing
            rst.ActiveConnection = "dsn=rule;uid=sa"
            rst.CursorLocation = adUseClient
            rst.CursorType = adOpenKeyset
            rst.LockType = adLockOptimistic
            rst.Source = "select rulename,note from rulepara where rulename='" & Trim$(List1.List(j)) & "'"
            rst.Open
            If s = "" Then
                If rst.RecordCount <> 0 And rst("rulename") <> "红签条" Then
                s = s + rst("note")
                End If
            Else
                If rst.RecordCount <> 0 And rst("rulename") <> "红签条" Then
                s = s + "," + rst("note")
                End If
            End If
            '判断是否是红铅条的任务
            If rst("rulename") = "红签条" Then
                s1 = " and note='紧急任务' "
            End If
    Next j
    
    Set rs1 = Nothing
    rs1.ActiveConnection = "dsn=dbw;uid=sa"
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenKeyset
    rs1.LockType = adLockOptimistic
    
    sql = "select * from t_myplantask where added='否'  "
    If s1 <> "" Then
        sql = sql & s1
    End If
    
    If s <> "" Then
        sql = sql & " order by " & s
    End If
    rs1.Source = sql
    rs1.Open

    
 
    If rs1.RecordCount <> 0 Then
         rs1.MoveFirst
         Do Until rs1.EOF
             Set rs2 = Nothing
             rs2.ActiveConnection = "dsn=dbw;uid=sa"
             rs2.CursorLocation = adUseClient
             rs2.CursorType = adOpenKeyset
             rs2.LockType = adLockOptimistic
             rs2.Source = "select machinenumber,status,timeoccupy from t_machineprocess,device  where t_machineprocess.drawingnumber='" & CStr(rs1("drawingno")) & _
                        "'" & "  and t_machineprocess.processnumber='" & CStr(rs1("processno")) & "'  and  t_machineprocess.machinenumber=device.deviceno"
             rs2.Open
             m = rs2.RecordCount
             If m <> 0 Then
                      ReDim machine(m)
                      ReDim machine1(m)
                      ReDim machine2(m)
                      rs2.MoveFirst
                      timeoccupy = rs2("timeoccupy")
                      quota = CSng(timeoccupy)
                      m = 1
                      Do Until rs2.EOF
                            Set rs3 = Nothing '提取每台设备的负荷,工作效率，及设备编号
                            rs3.ActiveConnection = "dsn=dbw;uid=sa"
                            rs3.CursorLocation = adUseClient
                            rs3.CursorType = adOpenKeyset
                            rs3.LockType = adLockOptimistic
                            rs3.Source = "select sum(timeoccupy) as number from t_submachineload where machinenumber='" & Trim$(rs2("machinenumber")) & "' group by machinenumber"
                            rs3.Open
                            
                            machine(m) = rs2("machinenumber") '提取设备编号
                            machine2(m) = rs2("status") '提取设备效率
                            If rs3.RecordCount <> 0 Then
                               rs3.MoveFirst
                              If rs3("number") <> "" Then
                                  machine1(m) = rs3("number") '提取设备负荷
                              Else
                                  machine1(m) = 0
                              End If
                            End If
                            m = m + 1
                            rs2.MoveNext
                     Loop
                 End If
                    Set rs3 = Nothing
                    Set rs2 = Nothing
                If m <> 0 Then
                Call mpop(rs1("planquantity"), quota, machine1(), machine2())    '调用过程进行任务分派
               
                m = UBound(machine())
                For i = 1 To m  '向设备负荷表里添加新数据
                    If C(i) <> 0 Then
                      rss.AddNew
                      rss("workcode") = CStr(rs1("workcode"))
                      rss("ordercode") = CStr(rs1("ordercode"))
                      rss("machinenumber") = CStr(machine(i))
                      rss("processnumber") = CStr(rs1("processno"))
                      rss("drawingnumber") = CStr(rs1("drawingno"))
                      rss("quantity") = CInt(C(i))
                      rss("timeoccupy") = (C(i) * quota) / machine2(i)
                      rss("plandate") = Year(Date) & "-" & Month(Date)
                      rss.Update
                    End If
                 Next i
                  Set rs3 = Nothing
                  rs3.ActiveConnection = "dsn=dbw;uid=sa"
                  rs3.CursorLocation = adUseClient
                  rs3.CursorType = adOpenKeyset
                  rs3.LockType = adLockOptimistic
                  rs3.Source = "t_subdaytaskplan"
                  rs3.Open
                 For i = 1 To m    '向任务表里添加新数据
                    If C(i) <> 0 Then
                     rs3.AddNew
                     rs3("workcode") = CStr(rs1("workcode"))
                     rs3("ordercode") = CStr(rs1("ordercode"))
                     rs3("drawingnumber") = CStr(rs1("drawingno"))
                     rs3("state") = CStr(rs1("processno"))
                     rs3("planquantity") = C(i)
                     rs3("machinecode") = CStr(machine(i))
                     rs3("playdate") = CDate(Date)
                     rs3.Update
                  End If
                 Next i
                 
            End If
            rs1.MoveNext
         Loop
       End If

     
  Set DataGrid3.DataSource = rss
  Call first(DataGrid3)
  Call initial(DataGrid3, "设备负荷表")
End Sub

Private Sub Comexit_Click()
  On Error Resume Next
   rs2.Close
   Set rs2 = Nothing
 Unload Me
End Sub

Private Sub Comexit2_Click()
  Unload Me
End Sub

Private Sub comshchjh_Click() '将大纲计划生成零件计划
  Dim i As Integer
  Dim dgrs As New ADODB.Recordset
  Dim sql As String
  Dim dgsql As String
  Dim rs3 As New ADODB.Recordset
  Screen.MousePointer = vbHourglass
  
 '2.大纲计划
           If myrs.State = adStateOpen Then myrs.Close
             myrs.ActiveConnection = "dsn=dbw;uid=sa"
             myrs.CursorLocation = adUseClient
             myrs.CursorType = adOpenKeyset
             myrs.LockType = adLockOptimistic
             myrs.Source = "DELETE FROM t_spbillofmaterials"
             myrs.Open
            If myrs.State = adStateOpen Then myrs.Close
             myrs.ActiveConnection = "dsn=dbw;uid=sa"
             myrs.CursorLocation = adUseClient
             myrs.CursorType = adOpenKeyset
             myrs.LockType = adLockOptimistic
             myrs.Source = "t_spbillofmaterials"
             myrs.Open
            
         On Error Resume Next
         If rs.RecordCount <> 0 Then '如果大纲记录不为零则生成零件记录
             rs.MoveFirst
             Do Until rs.EOF
                 On Error Resume Next
                 Call Main1(rs("workcode"), rs("ordercode"), rs("locomotivetype"), rs("drawingnumber"), rs("amount"))
                 If Not bool Then
                      rs("added") = "是"
                      rs.Update
                 End If
                 rs.MoveNext
              Loop
            If myrs.RecordCount <> 0 Then
               myrs.MoveFirst
               mark = myrs.Bookmark
               Do Until myrs.EOF
                    Call ff(myrs("workcode"), myrs("ordercode"), myrs("locomotivetype"), myrs("drawingnumber"), myrs("planquantity"), myrs("pargroupamount"))
                    myrs.Bookmark = mark
                    myrs.MoveNext
                    mark = myrs.Bookmark
               Loop
            End If
        End If
    '显示本月零部件计划
  
     sql = "select * from t_spgeneralpartplan" & _
         "  where added='否' order by ordercode"
        If rs2.State = adStateOpen Then rs2.Close
        rs2.ActiveConnection = "dsn=dbw;uid=sa"
        rs2.Source = sql
        rs2.CursorLocation = adUseClient
        rs2.CursorType = adOpenDynamic
        rs2.LockType = adLockOptimistic
        rs2.Open
        sql = "select * from t_myplantask" & _
         "  where added='否' order by ordercode"
        Set rs3 = Nothing
        rs3.ActiveConnection = "dsn=dbw;uid=sa"
        rs3.Source = sql
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenDynamic
        rs3.LockType = adLockOptimistic
        rs3.Open
        If rs3.RecordCount <> 0 Then rs3.MoveFirst
'        Set DataGrid4.DataSource = rs3
'        Call first(DataGrid4)
'        Call initial(DataGrid4, "工艺零件计划")
        If rs2.RecordCount <> 0 Then
         rs2.MoveFirst
        Do Until rs2.EOF
           Set dgrs = Nothing
           dgrs.ActiveConnection = "dsn=dbw;uid=sa"
           dgrs.CursorLocation = adUseClient
           dgrs.CursorType = adOpenKeyset
           dgrs.LockType = adLockOptimistic
           dgrs.Source = "select distinct processnumber  from t_subprocessplan where drawingnumber='" & Trim$(rs2("drawingnumber")) & "'"
           dgrs.Open
             If dgrs.RecordCount <> 0 Then
                dgrs.MoveFirst
                Do Until dgrs.EOF
                    rs3.AddNew
                    rs3("workcode") = CStr(rs2("workcode"))
                    rs3("ordercode") = CStr(rs2("ordercode"))
                    rs3("prodrawingnumber") = CStr(rs2("prodrawingnumber"))
                    rs3("drawingno") = CStr(rs2("drawingnumber"))
                    rs3("processno") = CStr(dgrs("processnumber"))
                    rs3("pargroupamount") = CInt(rs2("pargroupamount"))
                    rs3("planquantity") = CInt(rs2("planquantity"))
                    rs3("orderdate") = CDate(rs2("acceptdate"))
                    rs3("senddate") = CDate(rs2("senddate"))
                    rs3.Update
                    dgrs.MoveNext
                Loop
             End If
                rs2.MoveNext
       Loop
      End If
      Set rs3 = Nothing
'      DataGrid4.Visible = False
      Screen.MousePointer = vbDefault
        Set DataGrid2.DataSource = rs2
        Call initial(DataGrid2, "零件表")
        DataGrid2.Refresh
        Call first(DataGrid2)
    
comshchjh.Enabled = False

End Sub

Private Sub DTPicker1_Change()
    DataGrid2.Refresh
End Sub



Private Sub Form_Load()
 
 'Set conn = Nothing
 'conn.Open "dsn=dbw,UId=sa"
 Set rss = Nothing
  rss.ActiveConnection = "dsn=dbw;uid=sa"
  rss.CursorLocation = adUseClient
  rss.CursorType = adOpenKeyset
  rss.LockType = adLockOptimistic
  rss.Source = "t_submachineload"
  rss.Open
  Set DataGrid3.DataSource = rss
  Call first(DataGrid3)
  Call initial(DataGrid3, "设备负荷表")

 
 Dim sql As String
 DataGrid4.Visible = False
 sql = "select * from t_spgeneralpartplan where added='否'"
 
If rs2.State = adStateOpen Then rs2.Close
rs2.ActiveConnection = "dsn=dbw;uid=sa"
rs2.Source = sql
rs2.CursorLocation = adUseClient
rs2.CursorType = adOpenDynamic
rs2.LockType = adLockOptimistic
rs2.Open
Set DataGrid2.DataSource = rs2
    DataGrid2.Refresh
comshchjh.Enabled = False  '把生成零件计划按钮设成不可用
Call first(DataGrid2)
Call initial(DataGrid2, "零件表")
conn.ConnectionString = "DSN=rule;DATABASE=rule;;"
Set rsr = Nothing
rsr.ActiveConnection = "dsn=rule;uid=sa"
rsr.CursorLocation = adUseClient
rsr.CursorType = adOpenKeyset
rsr.LockType = adLockOptimistic
rsr.Source = "select * from rulepara"
rsr.Open

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Comexit_Click
   If myrs.State = adStateOpen Then myrs.Close
        myrs.ActiveConnection = "dsn=dbw;uid=sa"
        myrs.CursorLocation = adUseClient
        myrs.CursorType = adOpenKeyset
        myrs.LockType = adLockOptimistic
        myrs.Source = "DELETE FROM t_spbillofmaterials"
        myrs.Open

    End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call Comexit_Click
  If myrs.State = adStateOpen Then myrs.Close
  myrs.ActiveConnection = "dsn=dbw;uid=sa"
  myrs.CursorLocation = adUseClient
  myrs.CursorType = adOpenKeyset
  myrs.LockType = adLockOptimistic
  myrs.Source = "DELETE FROM t_spbillofmaterials"
  myrs.Open
 
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

  If SSTab1.TabIndex = 0 Then
        Set rs2 = Nothing
        rs2.ActiveConnection = "dsn=dbw;uid=sa"
       
        rs2.CursorLocation = adUseClient
        rs2.CursorType = adOpenDynamic
        rs2.LockType = adLockOptimistic
        rs2.Source = "select * from t_spgeneralpartplan where added='否'"
        rs2.Open
        Set DataGrid2.DataSource = rs2
            DataGrid2.Refresh
        Call first(DataGrid2)
        Call initial(DataGrid2, "零件表")
End If
cmbddgz.Clear
Dim i As Integer
rsr.MoveLast
rsr.MoveFirst
For i = 0 To rsr.RecordCount - 1
    cmbddgz.AddItem rsr.Fields("rulename"), i
    rsr.MoveNext
Next
cmbddgz.Text = cmbddgz.List(0)
End Sub

'drawingnumber为产品名称，quantitys为计划数量，pcocdssquatos为产品定额，manchine1 为设备的占用时间，manchine2为设备工作效率，
Public Sub mpop(quantitys As Integer, processquatos As Single, mach1() As Single, mach2() As Single)
    Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer
    Dim min, ave As Single
    m = UBound(mach1)
    ReDim C(m)
    For i = 1 To m
      C(i) = 0       '初始化每台设备的任务量为0
    Next i
       
   For i = 1 To quantitys '将n个任务依次分配到m太设备上
      k = 1
      min = mach1(1) + processquatos / mach2(1)
      For j = 2 To m       '找出设备占用时间最小的，并将任务分配给他
         ave = mach1(j) + processquatos / mach2(j)
         If (ave < min) Then
            min = ave
            k = j
         End If
       Next j
       mach1(k) = min
       C(k) = C(k) + 1
     Next i
End Sub
'根据已完工工序，获得未完工工序所需时间
Function GetRemainTime1(finishedprocess As String, drawno As String)
Dim tt As Integer
Dim rsf As New ADODB.Recordset
rsf.ActiveConnection = "dsn=dbw;uid=sa"
rsf.CursorLocation = adUseClient
rsf.CursorType = adOpenDynamic
rsf.LockType = adLockOptimistic
rsf.Source = "SELECT *  FROM t_machineprocess where drawingnumber='" & drawno & "'" & _
" and processnumber>=" & finishedprocess
rsf.Open
If rsf.RecordCount = 0 Then
    GetRemainTime1 = 0
Else
    tt = 0
    rsf.MoveFirst
    While Not rsf.EOF
        tt = rsf("timeoccupy") + tt
        rsf.MoveNext
    Wend
    GetRemainTime1 = tt
End If
rsf.Close

End Function
'根据已完工工序，获得未完工工序所需时间
Function GetTime1(drawno As String)
Dim tt As Integer
Dim rsf As New ADODB.Recordset
rsf.ActiveConnection = "dsn=dbw;uid=sa"
rsf.CursorLocation = adUseClient
rsf.CursorType = adOpenDynamic
rsf.LockType = adLockOptimistic
rsf.Source = "select *  from t_subprocessplan  where drawingnumber='" & drawno & "'"
rsf.Open
'若允许随机排序，须减去数组finishedprocess()中的工序
If rsf.RecordCount = 0 Then
    GetTime1 = 0
Else
    tt = 0
    rsf.MoveFirst
    While Not rsf.EOF
        tt = rsf("elapsetime") + tt
        rsf.MoveNext
    Wend
    GetTime1 = tt
End If
rsf.Close

End Function


