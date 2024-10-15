VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsuanfa1 
   Caption         =   "Form1"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10575
   ScaleWidth      =   13530
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   19288
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "调度算法"
      TabPicture(0)   =   "frmsuanfa1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "DataGrid1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "设备负荷情况查看"
      TabPicture(1)   =   "frmsuanfa1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DataGrid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Timer1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MSChart1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   8
         Top             =   600
         Width           =   14655
         Begin VB.ComboBox Cmbsuanfa 
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
            ItemData        =   "frmsuanfa1.frx":0038
            Left            =   1920
            List            =   "frmsuanfa1.frx":0042
            TabIndex        =   11
            Top             =   360
            Width           =   2415
         End
         Begin VB.CommandButton Comok 
            Caption         =   "任务分派"
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
            Left            =   6480
            TabIndex        =   10
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton comexit1 
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
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "算法"
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
            Left            =   1320
            TabIndex        =   13
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   7080
            TabIndex        =   12
            Top             =   240
            Width           =   90
         End
      End
      Begin VB.PictureBox MSChart1 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   13515
         TabIndex        =   7
         Top             =   9960
         Width           =   13575
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   720
         Top             =   7560
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Height          =   1095
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   14295
         Begin VB.ComboBox cmbmachine 
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
            ItemData        =   "frmsuanfa1.frx":0054
            Left            =   2160
            List            =   "frmsuanfa1.frx":0056
            TabIndex        =   5
            Top             =   480
            Width           =   2895
         End
         Begin VB.CommandButton combb 
            Caption         =   "报表"
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
            Left            =   9000
            TabIndex        =   4
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton Comexit 
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
            Left            =   11760
            TabIndex        =   3
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton Comfind 
            Caption         =   "查询"
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
            Left            =   6120
            TabIndex        =   2
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "设备选择"
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
            Left            =   1080
            TabIndex        =   6
            Top             =   600
            Width           =   840
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   7440
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   13123
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   8175
         Left            =   -74880
         TabIndex        =   15
         Top             =   1920
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14420
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
End
Attribute VB_Name = "frmsuanfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C() As Integer
Dim str1 As String
Dim rs As New ADODB.Recordset
Dim maa() As String
 Dim tt()
 Dim p()  As Integer '存储初始种群p(ss+1,nn+1)
 Dim pg() As Integer '存储新种群p(ss+1,nn+1)
 Dim A() As Integer '存储工艺相关数据a(mm+1,nn+1)
 Dim B() As Integer '存储工艺约束b（nn+1,ll+1)
 Dim pgg() As Integer '新产生的染色体串
 Dim D() As Integer '
 Dim dd As Integer
 Dim pnew() As Integer
 Dim fnew() As Single
Dim ran() As Single '存储随即数ran(nn+1)
Dim pran() As Single '存储各个染色体的选择概率pran(ss+1)
Dim pf() As Single '存储每个染色体的适值f(ss+1)
Dim f() As Single '存储每个染色体的适值f(ss+1)
Dim mach() As machine '用它来存储机器的任务集mach(mm+1,hh+1)
Dim mach1() As machine
Dim min() As Integer '用来记录设备当前的任务数

Dim ss As Integer '用它来设置种群大小
Dim pc As Single '用它来表示复制率
Dim pm As Single '用它来表示变异率
Dim mm As Integer '用它来表示机器数
Dim nn As Integer '用它来表示工序数
Dim ll As Integer '用它来表示约束工艺的最大数
Dim hh As Integer '用它来表示每台机器上的最多任务数

 
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
Sub AlgBuffFirst()
Dim sql As String, remaint As Integer
Dim finpro() As String
Set mrs = Nothing
sql = "select * from t_spgeneralpartplan order by drawingnumber"
mrs.Open sql, mconn, adOpenKeyset, adLockPessimistic
If mrs.RecordCount > 0 Then
    mrs.MoveFirst
    While Not mrs.EOF
        '获得已完工工序，对finpro()赋值
        ReDim finpro(1)
        finpro(0) = mrs("state")
        remaint = GetRemainTime(finpro, mrs("drawingnumber"))
        sql = mrs("senddate") - Date
        mrs("bufftime") = mrs("senddate") - Date - remaint
        mrs.Update
        mrs.MoveNext
    Wend
End If
mrs.Close
End Sub
Sub ljb()
Dim sql As String, remaint As Integer
Dim finpro() As String
Set mrs = Nothing
sql = "select * from t_spgeneralpartplan order by drawingnumber"
mrs.Open sql, mconn, adOpenKeyset, adLockPessimistic
If mrs.RecordCount > 0 Then
    mrs.MoveFirst
    While Not mrs.EOF
        '获得已完工工序，对finpro()赋值
        ReDim finpro(1)
        finpro(0) = mrs("state")
        remaint = GetRemainTime(finpro, mrs("drawingnumber"))
        sql = mrs("senddate") - Date
        If remaint <> 0 Then
            
        mrs("bufftime") = (mrs("senddate") - Date) / remaint
        End If
        mrs.Update
        mrs.MoveNext
    Wend
End If
mrs.Close
End Sub

Sub ShowResult()
Dim sql As String
Set mrs = Nothing
sql = "select drawingnumber as 图号,state as 下道工序,planquantity as 计划数量,senddate as 交货日期,"
Select Case Me.Tag
    Case "buff"
        sql = sql & " bufftime as 缓冲期 "
    Case "ljb"
        sql = sql & " bufftime as 临界比 "
    
End Select
 sql = sql & " from t_spgeneralpartplan order by bufftime"
mrs.Open sql, mconn, adOpenKeyset, adLockPessimistic
If mrs.RecordCount = 0 Then
    Set dgd_show.DataSource = Nothing
    dgd_show.Refresh
Else
    mrs.MoveFirst
    Set dgd_show.DataSource = mrs
    dgd_show.Refresh
End If
End Sub

'根据已完工工序，获得未完工工序所需时间
Function GetRemainTime(finishedprocess() As String, drawno As String)
Dim sql As String, tt As Integer
Dim rs As New ADODB.Recordset
sql = "select * from t_subpmreference where drawingnumber='" & drawno & "'"
'派工时工序固定从小到大，不能改变顺序
sql = sql & " and processnumber>=" & finishedprocess(0)
'若允许随机排序，须减去数组finishedprocess()中的工序
rs.CursorLocation = adUseClient
rs.Open sql, mconn, adOpenKeyset, adLockPessimistic
If rs.RecordCount = 0 Then
    GetRemainTime = 0
Else
    tt = 0
    rs.MoveFirst
    While Not rs.EOF
        tt = rs("elapsetime") + tt
        rs.MoveNext
    Wend
    GetRemainTime = tt
End If
rs.Close

End Function

Public Function findmachine(machine1 As String) As Integer
    Dim i As Integer
    Dim flag As Boolean
    flag = False
    i = 1
    Do Until i > mm Or flag
      If machine1 = maa(i) Then
          flag = True
      Else
        i = i + 1
      End If
    Loop
   If flag Then
    findmachine = i
   Else
    findmachine = 0
   End If
End Function
'总体初始化
Public Sub initial4()
    Dim i As Integer, j As Integer
    Dim findrs As New ADODB.Recordset
    Dim str As String
'ss = CInt(txtss.Text)
'   pc = CSng(txtpc.Text)
'  pm = CSng(txtpm.Text)
'  dd = CInt(txtdd.Text)
   ss = ss1
   pc = pc1
   pm = pm1
   dd = 4
'   dd = dd1
 
    '查询生产设备数
   Set rs = Nothing
    rs.ActiveConnection = "dsn=dbw;uid=sa"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Source = "select distinct machinenumber from t_machineprocess1,t_myplantask " & _
   " where added='否' and t_machineprocess1.drawingnumber=t_myplantask.drawingno " & _
   "and t_machineprocess1.processnumber=t_myplantask.processno order by machinenumber"
    rs.Open
    If rs.RecordCount = 0 Then
       MsgBox "当前没有加工设备", vbExclamation + vbInformation
       End
    End If
    mm = rs.RecordCount 'mm表示机器数
'    mm = CInt(txtmm.Text)
    ReDim maa(mm + 1)
    rs.MoveFirst
     i = 1
    Do Until rs.EOF
      maa(i) = Trim(rs("machinenumber"))
      i = i + 1
      rs.MoveNext
    Loop
    '求每台设备上的最大任务数
      hh = 0
     rs.MoveFirst
     Do Until rs.EOF
            Set findrs = Nothing
            findrs.ActiveConnection = "dsn=dbw;uid=sa"
            findrs.CursorLocation = adUseClient
            findrs.CursorType = adOpenKeyset
            findrs.LockType = adLockOptimistic
            findrs.Source = "select drawingno,processno " & _
              " from t_machineprocess1,t_myplantask " & _
              " where added='否' and t_machineprocess1.drawingnumber=t_myplantask.drawingno " & _
              "and t_machineprocess1.processnumber=t_myplantask.processno and t_machineprocess1.machinenumber='" & Trim$(rs("machinenumber")) & "'"
            findrs.Open
            If findrs.RecordCount > hh Then hh = findrs.RecordCount  'hh表示每台机器上的任务数
            rs.MoveNext
     Loop
    
    ll = 1

'   nn = CInt(txtnn.Text)
'   ll = CInt(txtll.Text)
'   hh = CInt(txthh.Text)
    '查询任务数
    Set rs = Nothing
    rs.ActiveConnection = "dsn=dbw;uid=sa"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Source = "select drawingno,processno,planquantity,timeoccupy ,machinenumber " & _
    "from t_myplantask ,t_machineprocess1 where t_machineprocess1.drawingnumber=t_myplantask.drawingno" & _
    " and t_machineprocess1.processnumber=t_myplantask.processno  and added='否' order by drawingno,processno"
    rs.Open
    If rs.RecordCount = 0 Then
       MsgBox "当前没有为排序的任务", vbExclamation + vbInformation
       End
    End If
    nn = rs.RecordCount
   ReDim p(ss + 1, nn + 1)
   ReDim pg(ss + 2, nn + 1)
   ReDim A(4, nn + 1)
   ReDim D(nn + 1)
   ReDim B(nn + 1, ll + 1)
   ReDim ran(nn + 1)
   ReDim f(ss + 1)
   ReDim mach(mm + 1, hh + 1)
   ReDim mach1(ss + 1, mm + 1, hh + 1)
   ReDim min(mm + 1)
   ReDim pf(ss + 1)
'   ReDim tt(1 To dd + 1)
   i = 1
   str = "ss"
   rs.MoveFirst
   Do Until rs.EOF
        A(1, i) = i   ' a存储工艺相关数据
        A(2, i) = findmachine(Trim$(rs("machinenumber")))
        A(3, i) = CSng(rs("timeoccupy") * rs("planquantity"))
        If str = Trim$(rs("drawingno")) Then
         B(i, 1) = i - 1
        Else
          B(i, 1) = 0
        End If
        str = Trim$(rs("drawingno"))
        i = i + 1
        rs.MoveNext
   Loop
 
End Sub
'遗传调度算法
Public Function genetic1() As Integer
  Dim k As Integer, i As Integer, j As Integer
  Dim imax As Single
  Dim sum As Single
  
  k = 1
  '初始化
  Call initial4
  '生成初始种群
  Call popsize(ss)
  Do Until k > dd
       '计算种群的适值
       Call translate(1)
        sum = 0
'        tt(k, 1) = k
        For i = 2 To ss
           sum = sum + 1 / f(i)
        Next i
'        tt(k) = sum / ss
       '计算选择概率
       Call fitness
       '选择操作
       Call chose
       '交叉操作
       Call crossover
       '变异操作
       Call mutation
       '重新设置种群
       Call fset
       '解码并计算适值
       'Call translate(1)
       k = k + 1
   Loop
   
   
   For i = 1 To ss
     '对每个染色体解码，
     Call ft(i, 1)
      For j = 1 To mm
         For k = 1 To hh
            mach1(i, j, k).gx = mach(j, k).gx
            mach1(i, j, k).start = mach(j, k).start
            mach1(i, j, k).stop = mach(j, k).stop
         Next k
      Next j
      imax = mach(1, min(1)).stop
              For j = 2 To mm
                If imax < mach(j, min(j)).stop Then imax = mach(j, min(j)).stop
              Next j
            f(i) = 1 / imax
    Next i
    k = 1
   ' 选出适值最大的染色体
   For i = 2 To ss
     If f(k) < f(i) Then
        k = i
     End If
   Next i
   genetic1 = k
  
End Function

'遗传禁忌调度算法
Public Function GA_TS() As Integer
  Dim k As Integer, i As Integer, j As Integer
  Dim imax As Single
  Dim sum As Single
  Dim tt(51) As Single
  
  k = 1
  '初始化
  Call initial4
  '生成初始种群
  Call popsize(ss)
  
  Do Until k > dd
       '计算种群的适值
       Call translate(1)
        sum = 0
'        tt(k, 1) = k
        For i = 2 To ss
           sum = sum + 1 / f(i)
        Next i
'        tt(k) = sum / ss
       '计算选择概率
       Call fitness
       '选择操作
       Call chose
       '交叉操作
       Call crossover
       '变异操作
       Call mutation
       '重新设置种群
       Call fset
    
        w = 1
   ' 选出适值最大的染色体
   For i = 2 To ss
     If f(w) < f(i) Then
        w = i
     End If
   Next i
   'tt(k) = 1 / f(w)
   'MSChart1.ChartData = tt
   'MSChart1.ColumnLabel = "完成时间"
   'MSChart1.RowLabel = "迭代次数"
   
   
   
    k = k + 1
   Loop
   For i = 1 To ss
     '对每个染色体解码，
     Call ft(i, 1)
      For j = 1 To mm
         For k = 1 To hh
            mach1(i, j, k).gx = mach(j, k).gx
            mach1(i, j, k).start = mach(j, k).start
            mach1(i, j, k).stop = mach(j, k).stop
         Next k
      Next j
      imax = mach(1, min(1)).stop
              For j = 2 To mm
                If imax < mach(j, min(j)).stop Then imax = mach(j, min(j)).stop
              Next j
            f(i) = 1 / imax
    Next i
    k = 1
   ' 选出适值最大的染色体
   For i = 2 To ss
     If f(k) < f(i) Then
        k = i
     End If
   Next i
   GA_TS = k
   

  
End Function

'重组算子
Public Sub tsm()
Dim i As Integer
Dim sum As Double  '适配值和
Dim ave As Single '适配值均值
Dim Tabu(50, 50) As Single '禁忌表
Dim flag As Boolean
Dim f1 As Single '父染色体一的适值
Dim f2 As Single '父染色体二的适值
Dim f3 As Single
Dim fn As Single '后代染色体的适值
Dim f(50) As Single '存储每个染色体的适值f(ss+1)

flag = False
sum = 0
For i = 1 To ss
  sum = sum + f(i)
Next i
ave = sum / ss
If f3 < ave Then
  ss = ss + 1
  Call tabu_change
Else
  For i = 1 To ss
    If Tabu(i, 1) = f3 Then
       flag = True
    End If
   Next i
  If flag = False Then
    ss = ss + 1
    Call tabu_change
    End If
End If
If flag = True Then
   fn = ddd(f1, f2, f3)
        If fn = f1 And fn <> f3 Then
            For t = 1 To nn
               If pg(ss + 1, t) = 0 Then Stop
               pg(h1, t) = pg(ss + 1, t)
            Next t
        ElseIf fn = f2 And fn <> f3 Then
           For t = 1 To nn
               If pg(ss + 1, t) = 0 Then Stop
               pg(h2, t) = pg(ss + 1, t)
            Next t
       End If
End If
For t = 1 To nn
     pg(ss + 1, t) = 0
Next t
End Sub


'禁忌表的移动
Public Sub tabu_change()
Dim Tabu(100, 100) As Single '禁忌表
Dim i As Integer
For i = 1 To ss
 If Tabu(i, 1) = f3 And Tabu(i, 2) <= 3 Then
    Tabu(i, 2) = Tabu(i, 2) - 1
 Else
   If Tabu(i, 0) = 0 Then
      Tabu(i, 1) = f3
      Tabu(i, 2) = 3
    End If
  End If
Next i
End Sub
Public Sub fset()
  Dim i As Integer, j As Integer
  For i = 1 To ss
    For j = 1 To nn
    p(i, j) = pg(i, j)
    Next j
  Next i
  
End Sub
'本函数用来计算各个染色体的选择概率
Public Sub fitness()
    Dim i As Integer, sum As Single
    Dim ppp() As Single
    ReDim ppp(ss)
    sum = 0
    '计算染色体的适值和
    For i = 1 To ss
      sum = sum + f(i)
    Next i
    '计算染色体的选择概率
    For i = 1 To ss
       ppp(i) = f(i) / sum
    Next i
    '计算割染色体的累积概率pf
      pf(0) = 0
    For i = 1 To ss
       pf(i) = pf(i - 1) + ppp(i)
    Next i
End Sub
Public Function ddd(A As Single, B As Single, C As Single) As Single
    Dim k As Single
    k = A
    If k > B Then
      k = B
    End If
    If k > C Then
       k = C
    End If
    ddd = k
    
End Function
'该函数返回机器生产时间
Public Function fit(g As Integer, str As Integer) As Single
     Dim k As Integer, j As Integer
       Call ft(g, str)
              imax = mach(1, min(1)).stop
              For j = 2 To mm
                If imax < mach(j, min(j)).stop Then imax = mach(j, min(j)).stop
              Next j
            fit = 1# / imax
End Function
'该函数返回本条染色体的适值
Public Function fit1(g As Integer, str As Integer) As Single
     Call ft1(g, str)
     fit1 = 1 / fmax()
End Function
'编写选择函数
Public Sub chose()
    Dim i As Integer, j As Integer, k As Single
    Dim flag As Boolean
    Dim C As Integer
     Randomize
    For i = 1 To ss
        k = Rnd()
        j = 1
        flag = False
        Do Until flag Or j > ss
           If k > pf(j - 1) And k < pf(j) Then
              flag = True
           Else
             j = j + 1
           End If
        Loop
        
          For C = 1 To nn
            If p(j, C) = 0 Then Stop
             pg(i, C) = p(j, C)
          Next C
       
      Next i
End Sub
'在某串基因串中查找某基因的后继,参数g表明第g条染色体，n和m表明从n到m的基因串，l为要查找的基因
Public Function findhj(g As Integer, n As Integer, m As Integer, l As Integer) As Boolean
     Dim i As Integer, k As Integer, j As Integer
       findhj = False
     For i = n To m
       k = pg(g, i)
       j = 1
       Do Until B(k, j) = 0 Or findhj
           If B(k, j) = l Then
              findhj = True
           End If
           j = j + 1
       Loop
     Next i
End Function

'在某串基因串中查找某基因的后继,参数g表明第g条染色体，n和m表明从n到m的基因串，l为要查找的基因
Public Function findqq(g As Integer, n As Integer, m As Integer, l As Integer) As Boolean
     Dim i As Integer, k As Integer
      If B(l, 1) = 0 Then
         findqq = False
      Else
         findqq = False
         k = 1
         Do Until B(l, k) = 0 Or findqq
            findqq = find(g, n, m, B(l, k))
            k = k + 1
         Loop
     End If
End Function

'在某段基因串中查找某个基因是否存在
'染色体中从n到m的基因串，查找基因l,g代表地g条染色体
Public Function find(g As Integer, n As Integer, m As Integer, l As Integer) As Boolean
     Dim i As Integer
     find = False
     i = n
     Do Until i > m Or find
     If l = pg(g, i) Then
        find = True
     End If
        i = i + 1
     Loop
End Function
'在某段基因串中查找某个基因是否存在
'染色体中从n到m的基因串，查找基因l
Public Function finddd(n As Integer, m As Integer, l As Integer) As Boolean
     Dim i As Integer
     find = False
     i = n
     Do Until i > m Or find
     If l = pgg(i) Then
        find = True
     End If
        i = i + 1
     Loop
End Function
'变异操作
Public Sub mutation()
 Dim fran() As Single
 Dim i As Integer, j As Integer, k As Integer, r As Integer
 Dim yy As Integer, ww As Integer, l As Integer
 Dim cro() As Single
 Dim flag As Boolean
 Dim jj As Integer
 Dim f1 As Single, f2 As Single
 Randomize
 'k为应发生变异的基因的个数
 k = pm * ss * nn
 ReDim cro(k + 1)
 ReDim fran(ss * nn + 1)
 '生成ss*nn个随机数
 For i = 1 To ss * nn
    fran(i) = Rnd
 Next i
 j = 0
 i = 1
 flag = True
 Do Until j >= k Or i > ss * nn
   If fran(i) < pm Then
      j = j + 1
      cro(j) = i
   End If
   i = i + 1
 Loop
 '发生移位变异
 For r = 1 To j
   'yy表示变异的染色体ww表示变异的基因
        yy = cro(r) \ nn + 1
        ww = cro(r) Mod (nn)
        If ww = 0 Then
           yy = yy - 1
           ww = nn
        End If
        'l为基因移位的位置
        l = Int((nn * Rnd) + 1)
        For i = 1 To nn
        pg(ss + 1, i) = pg(yy, i)
        Next i
        jj = pg(ss + 1, ww)
        If l < ww Then
             '在基因串pg(yy,l)到pg(yy,ww-1)中查找基因 pg(yy, ww)的前驱
             If findqq(ss + 1, l, ww - 1, jj) Then
                 
             Else
                
                For i = ww - 1 To l Step -1
                    pg(ss + 1, i + 1) = pg(ss + 1, i)
                Next i
                 pg(ss + 1, l) = jj
             End If
        ElseIf l > ww Then
             '在基因串pg(yy,l)到pg(yy,ww-1)中查找基因 pg(yy, ww)的后记
             If findhj(yy, ww + 1, l, jj) Then
                 
             Else
                For i = ww + 1 To l
                    pg(ss + 1, i - 1) = pg(ss + 1, i)
                Next i
                 pg(ss + 1, l) = jj
             End If
        End If
        f1 = fit(yy, 2)
        f2 = fit(ss + 1, 2)
        If f2 > f1 Then
           For i = 1 To nn
           pg(yy, i) = pg(ss + 1, i)
           Next i
        End If
 Next r
 
End Sub
Public Sub translate1(str As Integer)
   Dim i As Integer, j As Integer
   
     '初始化适值
      For j = 1 To ss
         f(j) = 0
      Next j
   '解码
       For i = 1 To ss
              'ft1为解码函数，fmax 为求最大机器加工时间函数
              Call ft1(i, str)
              f(i) = 1 / fmax()
       Next i
End Sub

Public Sub mutation1()
 Dim fran() As Single
 Dim i As Integer, j As Integer, k As Integer, r As Integer, h As Integer
 Dim yy As Integer, ww As Integer, l As Single
 Dim cro() As Single
 Dim jj As Integer
 ReDim pnew(mm + 1, nn + 1)
 ReDim fnew(mm + 1)
 Randomize
 'k为应发生变异的基因的个数
 k = pm * ss * nn
 ReDim cro(k + 1)
 ReDim fran(ss * nn + 1)
 '生成ss*nn个随机数
 For i = 1 To ss * nn
    fran(i) = Rnd
 Next i
 j = 0
 i = 1
 '查找要发生变异的染色体号
 Do Until j >= k Or i > ss * nn
   If fran(i) < pm Then
      j = j + 1
      cro(j) = i
   End If
   i = i + 1
 Loop
 '发生插入变异
 
 For i = 1 To j
     yy = cro(i) \ nn + 1
     ww = cro(i) Mod nn
     If ww = 0 Then
       yy = yy - 1
       ww = nn
     End If
     For h = 1 To pg(yy, ww) - 1
            For r = 1 To ww - 1
              
              pnew(h, r) = pg(yy, r)
              If pnew(h, r) = 0 Then Stop
            Next r
        
             pnew(h, ww) = h
            For r = ww + 1 To nn
              pnew(h, r) = pg(yy, r)
              If pnew(h, r) = 0 Then Stop
            Next r
           fnew(h) = fit1(h, 3)
      Next h
      For h = pg(yy, ww) + 1 To mm
            For r = 1 To ww - 1
             
              pnew(h, r) = pg(yy, r)
              If pnew(h, r) = 0 Then Stop
            Next r
            pnew(h, ww) = h
            For r = ww + 1 To nn
            
              pnew(h, r) = pg(yy, r)
              If pnew(h, r) = 0 Then Stop
            Next r
            fnew(h) = fit1(h, 3)
      Next h
            fnew(pg(yy, ww)) = fit1(yy, 2)
     '进行局部搜索
       l = fmax1(mm)
       If l = pg(yy, ww) Then
       Else
         For r = 1 To nn
           pg(yy, r) = pnew(l, r)
         Next r
       End If
 Next i

End Sub
'交叉操作 采用lox法线性顺序交叉
Public Sub crossover1()
   Dim cro() As Integer
   Dim i As Integer, j As Integer, k As Integer
   Dim flag As Boolean, flag1 As Boolean
'   Dim f1 As Single '父染色体一的适值
'   Dim f2 As Single '父染色体二的适值
'   Dim fh1 As Single, fh2 As Single '后代染色体的适值
   Dim h1 As Integer '交叉染色体的号码
   Dim h2 As Integer '交叉染色体的号码
   Dim g1 As Integer '断点一
   Dim g2 As Integer '断点二
   Dim t As Integer
  Dim w1 As Single, w2 As Single, w3 As Single, w4 As Single
   k = ss * pc '它记录了发生交叉的染色体的个数
   ReDim pran(ss) '它记录了每条染色体的交叉随机数
   ReDim cro(k + 1) As Integer
   Randomize
   '生成每条染色体地交叉随机数
   For i = 1 To ss
      pran(i) = Rnd
   Next i
      j = 0
      i = 1
    '查找要发生交叉地染色体地号码
   Do Until i > ss Or j >= k
       If pran(i) < pc Then
          j = j + 1
           cro(j) = i
       End If
       i = i + 1
   Loop
   
   i = 2
   '染色体交叉
   k = j
  Do Until i > k
     h1 = cro(i - 1)
     h2 = cro(i)
     g1 = Int((nn * Rnd) + 1) '产生一个1到nn的随机数
     g2 = Int((nn * Rnd) + 1)
      If g1 > g2 Then
         t = g1
         g1 = g2
         g2 = t
      End If
      '顺序交叉
      
      For j = 1 To g1 - 1
          pg(ss + 1, j) = pg(h1, j)
          pg(ss + 2, j) = pg(h2, j)
      Next j
      For j = g1 To g2
          pg(ss + 1, j) = pg(h2, j)
          pg(ss + 2, j) = pg(h1, j)
      Next j
       For j = g2 + 1 To nn
          pg(ss + 1, j) = pg(h1, j)
          pg(ss + 2, j) = pg(h2, j)
      Next j
      '计算双亲和后代染色体的适值
    
  w1 = fit1(h1, 2)
  w2 = fit1(h2, 2)
  w3 = fit1(ss + 1, 2)
  w4 = fit1(ss + 2, 2)
   If w3 > w1 Or w3 > w2 Then
      If w1 > w2 Then
         For j = 1 To nn
            pg(h2, j) = pg(ss + 1, j)
         Next j
      Else
         For j = 1 To nn
           pg(h1, j) = pg(ss + 1, j)
         Next j
      End If
   End If
   If w4 > w1 Or w3 > w2 Then
      If w1 > w2 Then
         For j = 1 To nn
           pg(h2, j) = pg(ss + 2, j)
         Next j
      Else
         For j = 1 To nn
           pg(h1, j) = pg(ss + 2, j)
         Next j
      End If
   End If

          i = i + 2
  Loop
  End Sub
'交叉操作 采用lox法线性顺序交叉
Public Sub crossover()
   Dim cro() As Integer
   Dim i As Integer, j As Integer, k As Integer
   Dim flag As Boolean, flag1 As Boolean
   Dim f1 As Single '父染色体一的适值
   Dim f2 As Single '父染色体二的适值
   Dim fn As Single '后代染色体的适值
   Dim h1 As Integer '交叉染色体的号码
   Dim h2 As Integer '交叉染色体的号码
   Dim g1 As Integer '断点一
   Dim g2 As Integer '断点二
   Dim t As Integer
   Dim www As Integer
   Dim f3 As Single
   k = ss * pc '它记录了发生交叉的染色体的个数
   ReDim pran(ss) '它记录了每条染色体的交叉随机数
   ReDim cro(k + 1) As Integer
   Randomize
   For i = 1 To ss
      pran(i) = Rnd
   Next i
      j = 0
'      flag1 = False
      i = 1
   Do Until i > ss Or j >= k
       If pran(i) < pc Then
          j = j + 1
           cro(j) = i
       End If
       i = i + 1
   Loop
   
   i = 2
   '染色体交叉
   k = j
  Do Until i > k
     h1 = cro(i - 1)
     h2 = cro(i)
     g1 = Int((nn * Rnd) + 1) '产生一个1到nn的随机数
     g2 = Int((nn * Rnd) + 1)
      If g1 > g2 Then
         t = g1
         g1 = g2
         g2 = t
      End If
      '顺序交叉
      Dim w1 As Integer
      Dim w2 As Integer
      For j = g1 To g2
          pg(ss + 1, j) = pg(h1, j)
      Next j
      w1 = g1
      w2 = g2
      www = 0
    For t = 1 To nn
       '看基因pg(h2, t)是否在染色体中
       If Not find(ss + 1, w1, w2, pg(h2, t)) Then
          If findqq(ss + 1, w1, w2, pg(h2, t)) Then
                w2 = w2 + 1
                If w2 > nn Then
                  For j = w1 To w2 - 1
                     pg(ss + 1, j - 1) = pg(ss + 1, j)
                  Next j
                  w1 = w1 - 1
                  w2 = w2 - 1
                  pg(ss + 1, w2) = pg(h2, t)
                Else
                 pg(ss + 1, w2) = pg(h2, t)
                End If
          Else
             www = www + 1
                If www >= w1 Then
                    If findhj(ss + 1, w1, w2, pg(h2, t)) Then
                         For j = w2 To w1 Step -1
                          pg(ss + 1, j + 1) = pg(ss + 1, j)
                         Next j
                         w2 = w2 + 1
                         w1 = w1 + 1
                         pg(ss + 1, www) = pg(h2, t)
                    Else
                     w2 = w2 + 1
                     pg(ss + 1, w2) = pg(h2, t)
                     www = www - 1
                    End If

                Else
                 pg(ss + 1, www) = pg(h2, t)
                End If
          End If
       End If
    Next t
        f1 = fit(h1, 2)
        f2 = fit(h2, 2)
        f3 = fit(ss + 1, 2)
        fn = ddd(f1, f2, f3)
        If fn = f1 And fn <> f3 Then
            For t = 1 To nn
               If pg(ss + 1, t) = 0 Then Stop
               pg(h1, t) = pg(ss + 1, t)
            Next t
        ElseIf fn = f2 And fn <> f3 Then
           For t = 1 To nn
               If pg(ss + 1, t) = 0 Then Stop
               pg(h2, t) = pg(ss + 1, t)
            Next t
       End If
           For t = 1 To nn
              pg(ss + 1, t) = 0
           Next t
       i = i + 2
  Loop
End Sub
'此函数为解码函数
'根据染色体解码,参数n为第几条染色体，str代表初始种群还是新生种群
Public Sub ft(n As Integer, str As Integer)
    Dim i As Integer, j As Integer, k As Integer, h As Integer
    Dim m As Integer '代表工序所对应的机器号
'   Dim m1 As Integer'代表第一台机器的第几到工序
'   Dim m2 As Integer'代表第二台机器的第几到工序
'   Dim m3 As Integer '代表第三台机器的第几到工序
    Dim pre1 As Single
    Dim flag As Boolean
    Dim lg As Single
    '初始化任务量为0
    
        For j = 1 To mm
            min(j) = 0
        Next j
   
        For j = 1 To mm
         For k = 1 To hh
          mach(j, k).gx = 0
           mach(j, k).start = 0
           mach(j, k).stop = 0
         Next k
      Next j
    Select Case str
    Case 1
        For j = 1 To nn
                  '求出机器号
                    m = A(2, p(n, j))
 
                    min(m) = min(m) + 1
                    mach(m, min(m)).gx = p(n, j)
                    pre1 = precede(p(n, j))
                    If mach(m, min(m) - 1).stop < pre1 Then
                        mach(m, min(m)).start = pre1
                        mach(m, min(m)).stop = mach(m, min(m)).start + A(3, p(n, j))
                    Else
                    '查找本道工序应插入的最好地方
                       flag = False
                       k = 1
                      Do Until flag Or k > min(m)
                         If mach(m, k).start > pre1 And mach(m, k).start - mach(m, k - 1).stop > A(3, p(n, j)) Then
                             If pre1 < mach(m, k - 1).stop Then
                                flag = True
                             ElseIf mach(m, k).start - pre1 > A(3, p(n, j)) Then
                                flag = True
                             Else
                                k = k + 1
                             End If
                         Else
                            k = k + 1
                         End If
                      Loop
                      '如果找到插入点则作如下操作
                      If Not flag Then
                         mach(m, min(m)).start = mach(m, min(m) - 1).stop
                         mach(m, min(m)).stop = mach(m, min(m)).start + A(3, p(n, j))
                      Else
                         For h = min(m) To k Step -1
                             mach(m, h).gx = mach(m, h - 1).gx
                             mach(m, h).start = mach(m, h - 1).start
                             mach(m, h).stop = mach(m, h - 1).stop
                         Next h
                              mach(m, k).gx = p(n, j)
                             If pre1 > mach(m, k - 1).stop Then
                                mach(m, k).start = per1
                             Else
                                 mach(m, k).start = mach(m, k - 1).stop
                             End If
                             mach(m, k).stop = mach(m, k).start + A(3, p(n, j))
                      End If
                   End If
                    
            Next j
     Case 2
             For j = 1 To nn
                  '求出机器号
                    m = A(2, pg(n, j))
                    min(m) = min(m) + 1
                    mach(m, min(m)).gx = pg(n, j)
                    pre1 = precede(pg(n, j))
                    If mach(m, min(m) - 1).stop < pre1 Then
                        mach(m, min(m)).start = pre1
                        mach(m, min(m)).stop = mach(m, min(m)).start + A(3, pg(n, j))
                    Else
                    '查找本道工序应插入的最好地方
                       flag = False
                       k = 1
                      Do Until flag Or k > min(m)
                         If mach(m, k).start > pre1 And mach(m, k).start - mach(m, k - 1).stop > A(3, pg(n, j)) Then
                             If pre1 < mach(m, k - 1).stop Then
                                flag = True
                             ElseIf mach(m, k).start - pre1 > A(3, pg(n, j)) Then
                                flag = True
                             Else
                                k = k + 1
                             End If
                         Else
                            k = k + 1
                         End If
                      Loop
                      '如果找到插入点则作如下操作
                      If Not flag Then
                         mach(m, min(m)).start = mach(m, min(m) - 1).stop
                         mach(m, min(m)).stop = mach(m, min(m)).start + A(3, pg(n, j))
                      Else
                         For h = min(m) To k Step -1
                             mach(m, h).gx = mach(m, h - 1).gx
                             mach(m, h).start = mach(m, h - 1).start
                             mach(m, h).stop = mach(m, h - 1).stop
                         Next h
                             If pre1 > mach(m, k - 1).stop Then
                                mach(m, k).start = per1
                             Else
                                 mach(m, k).start = mach(m, k - 1).stop
                             End If
                             mach(m, k).stop = mach(m, k).start + A(3, pg(n, j))
                      End If
                   End If
                    
            Next j

     End Select

End Sub
'此函数为解码函数
'根据染色体解码,参数n为第几条染色体，str代表初始种群还是新生种群
Public Sub ft1(n As Integer, str As Integer)
    Dim j As Integer
    Dim m As Integer '代表工序所对应的机器号
     '初始化设备加工时间
        For j = 1 To mm
            min(j) = 0
        Next j
    Select Case str
    Case 1
        For j = 1 To nn
            '求出机器号
            m = p(n, j)
            min(m) = min(m) + A(j, m)
        Next j
     Case 2
             For j = 1 To nn
                    '求出机器号
                    m = pg(n, j)
                    min(m) = min(m) + A(j, m)
            Next j
     Case 3
          For j = 1 To nn
                    '求出机器号
                    m = pnew(n, j)
                    
                    min(m) = min(m) + A(j, m)
                    If min(m) = 0 Then Stop
          Next j
     End Select

End Sub
'求最大的机器加工时间
Public Function fmax() As Single
     Dim i As Integer
     fmax = min(1)
     For i = 2 To mm
        If fmax < min(i) Then
           fmax = min(i)
        End If
     Next i
End Function
'求最大的机器加工时间
Public Function fmax1(j As Integer) As Single
     Dim i As Integer
     fmax1 = 1
     For i = 2 To j
        If fnew(fmax1) < fnew(i) Then
           fmax1 = i
        End If
     Next i
End Function

' 此函数为计算适值函数，他将染色体翻译成解,并求出解的适值
Public Sub translate(str As Integer)
   Dim i As Integer, j As Integer, k As Integer, h As Integer
   Dim m As Integer '代表工序所对应的机器号
   Dim imax As Single
   Dim lg As Single
     '初始化适值
      For j = 1 To ss
         f(j) = 0
      Next j
      '初始化
'      For j = 1 To mm
'         For k = 1 To hh
'          mach(j, k).gx = 0
'           mach(j, k).start = 0
'           mach(j, k).stop = 0
'         Next k
'      Next j
   '解码
       For i = 1 To ss
              'ft为解码函数
              Call ft(i, str)
              imax = mach(1, min(1)).stop
              For j = 2 To mm
                If imax < mach(j, min(j)).stop Then imax = mach(j, min(j)).stop
              Next j
            f(i) = 1 / imax
       Next i
End Sub
'参数n代表工序号，返回本道工序的约束工序的最迟完成时间
'本函数来求工序n的约束工序的最迟完成时间
Public Function precede(n As Integer) As Single
    Dim i As Integer, k As Integer, j As Integer
    i = 1
    precede = 0
    'b(n,i)为工序n的约束工艺
    Do Until i > ll Or B(n, i) = 0
     '找出生产b(n,i)工序的机器号k
       k = A(2, B(n, i))
         For j = min(k) To 1 Step -1
            If mach(k, j).gx = B(n, i) Then
              If precede < mach(k, j).stop Then
                precede = mach(k, j).stop
              End If
            End If
         Next j
              i = i + 1
   Loop
    
End Function

'参数说明a为代表工艺、加工机器，及加工施加的数组，其大小为a(n,3),
'b为表示工艺间约束关系的二维数组，它为b(n,m);
'size设定初始种群的大小，n即可代表共艺数，又可代表染色体长度
Public Sub popsize(size As Integer)
   Dim s() As Integer '表示当前可调度工艺
   Dim i As Integer
   Dim k As Integer
   Dim h As Integer
   Dim j As Integer
   Dim imax As Integer
   ReDim s(nn)
   Randomize
   
 '生成初始种群
  For h = 1 To size
   '初始化d(n)
   
    i = 1
    Do Until i > nn
       If B(i, 1) = 0 Then
          D(i) = 0
       Else
          D(i) = 1
       End If
       i = i + 1
    Loop
   '初始化s(n)
    i = 1
    k = 1
   Do Until i > nn
      If D(i) = 0 Then
         s(k) = i
         D(i) = 2
         k = k + 1
       End If
       i = i + 1
   Loop
       k = k - 1
   '生成初始种群中的一条染色体
   For j = 1 To nn
       '计算k割可调度工序的随机数
        For i = 1 To k
           ran(i) = Rnd()
        Next i
       'imax为可放入染色体中工序的序号
         imax = big(k)
         p(h, j) = s(imax)
         If k <> imax Then
            s(imax) = s(k)
         End If
        
         '设置每个工序的约束状态
         Call change(h, j)
          '设置可调度工艺序列s
         i = 1
         Do Until i > nn
           If D(i) = 0 Then
                s(k) = i
                D(i) = 2
                k = k + 1
           End If
           i = i + 1
         Loop
         k = k - 1
    Next j
 Next h

End Sub


'参数说明a为代表工艺、加工机器，及加工施加的数组，其大小为a(n,3),
'b为表示工艺间约束关系的二维数组，它为b(n,m);
'size设定初始种群的大小，n即可代表共艺数，又可代表染色体长度
Public Sub popsize1(size As Integer)
   Dim h As Integer
   Dim j As Integer
   Randomize
 '生成初始种群
  For h = 1 To size
     '生成初始种群中的一条染色体
     For j = 1 To nn
       '生成一个1到mm的随机数
         p(h, j) = Int(mm * Rnd) + 1
       'imax为可放入染色体中工序的序号
       Next j
 Next h
End Sub
'用来改变每道工艺的约素状态
'参数n代表第几条染色体m代表当前染色体的长度
Public Sub change(n As Integer, m As Integer)
   Dim flag1 As Boolean, flag2 As Boolean
   Dim i As Integer, j As Integer, k As Integer
     k = 1
   Do Until k > nn
    If D(k) = 1 Then
           flag1 = False
           i = 1
        Do Until i > ll Or flag1 Or B(k, i) = 0
             j = 1
             flag2 = False
                Do Until (j > m Or flag2)
                   If p(n, j) = B(k, i) Then
                      flag2 = True
                   End If
                      j = j + 1
                Loop
                i = i + 1
            If B(k, i) = 0 Then
               If flag2 Then
                   flag1 = True
               End If
            End If
        Loop
         If flag1 Then
           D(k) = 0
         End If
    End If
       k = k + 1
 Loop
End Sub
'参数k表示随机数的个数
Public Function big(k As Integer) As Integer
     Dim i As Integer
      big = 1
      For i = 2 To k
        If ran(big) < ran(i) Then
            big = i
        End If
      Next i
End Function

Private Sub Cmbsuanfa_LostFocus()
    If Trim$(Cmbsuanfa.Text) = "算法1" Then
        Comok.Caption = "参数设置"
        str1 = "算法1"
    ElseIf Trim$(Cmbsuanfa.Text) = "算法2" Then
            Comok.Caption = "参数设置"
            str1 = "算法2"
         
    End If
    
End Sub

Private Sub Combb_Click()
   Dim bbrs As New ADODB.Recordset
  Dim findrs As New ADODB.Recordset
  If rs.RecordCount <> 0 Then
        Set bbrs = Nothing
        bbrs.ActiveConnection = "dsn=dbw;uid=sa"
        bbrs.CursorLocation = adUseClient
        bbrs.CursorType = adOpenKeyset
        bbrs.LockType = adLockOptimistic
        bbrs.Source = "DELETE FROM t_machine"
        bbrs.Open
        Set bbrs = Nothing
        bbrs.ActiveConnection = "dsn=dbw;uid=sa"
        bbrs.CursorLocation = adUseClient
        bbrs.CursorType = adOpenKeyset
        bbrs.LockType = adLockOptimistic
        bbrs.Source = "select * from  t_machine  "
        bbrs.Open
        rs.MoveFirst
      Do Until rs.EOF
         bbrs.AddNew
         bbrs("ordercode") = CStr(rs("ordercode"))
         bbrs("workcode") = CStr(rs("workcode"))
         bbrs("machinecode") = CStr(rs("machinenumber"))
         Set findrs = Nothing
        findrs.ActiveConnection = "dsn=dbw;uid=sa"
        findrs.CursorLocation = adUseClient
        findrs.CursorType = adOpenKeyset
        findrs.LockType = adLockOptimistic
        findrs.Source = "select devicename from  device  where deviceno='" & Trim$(rs("machinenumber")) & "'"
        findrs.Open
        If findrs.RecordCount <> 0 Then
          findrs.MoveFirst
          bbrs("machinename") = CStr(findrs("devicename"))
         End If
         bbrs("drawingnumber") = CStr(rs("drawingnumber"))
         bbrs("quantity") = CInt(rs("quantity"))
         bbrs("state") = CInt(rs("processnumber"))
         bbrs("timeoccupy") = CSng(rs("timeoccupy"))
         bbrs.Update
         rs.MoveNext
      Loop
  End If
      On Error Resume Next
     'CrystalReport1.ReportFileName = App.Path & "\report\machine.rpt"
     'CrystalReport1.Action = 1
        
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Comexit_Click()
  Unload Me
End Sub

Private Sub Comexit1_Click()
 Unload Me
End Sub

Private Sub comfind_Click()
   Dim number As Integer
   Dim str As String
   If cmbmachine.Text = "" Then
      MsgBox "请您选择设备", vbExclamation + vbInformation
      Exit Sub
   End If
    number = InStr(1, cmbmachine.Text, "/", vbTextCompare)
    str = Left(Trim$(cmbmachine.Text), number - 1)
    Set rs = Nothing
    rs.ActiveConnection = "dsn=dbw;uid=sa"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Source = "select * from t_submachineload where machinenumber='" & str & "'"
    rs.Open
    If rs.RecordCount = 0 Then
       MsgBox "此设备上目前没有派任务", vbExclamation + vbInformation
    End If
    combb.Enabled = True
    Set DataGrid2.DataSource = rs
    Call first(DataGrid2)
    Call initial(DataGrid2, "设备负荷表")
    
End Sub

Private Sub comfresh_Click()
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

Private Sub Comgtt_Click()
Dim findsql As String
Dim findrs As New ADODB.Recordset
If Trim$(cmbmachine1.Text) = "" Then
   MsgBox "请您选择查看的设备类", vbExclamation + vbInformation
   Exit Sub
End If
If Option1.Value = True Then
   MSChart2.chartType = VtChChartType2dLine
Else
   MSChart2.chartType = VtChChartType2dBar
End If
'findsql = "select  productname,totalamount,max(stateamount) as amount,state   from t_mmmobilestock where drawingnumber='" & Trim$(cmbproductname) & "' group by productname,totalamount,state"
findsql = "select sum(timeoccupy) as occupy,machinenumber from t_submachineload,device where machinenumber=deviceno and devicename='" & cmbmachine1.Text & "' group by machinenumber order by machinenumber"
Set findrs = Nothing
findrs.ActiveConnection = "dsn=dbw;uid=sa"
findrs.CursorLocation = adUseClient
findrs.CursorType = adOpenKeyset
findrs.LockType = adLockOptimistic
findrs.Source = findsql
findrs.Open

Dim sum As Integer
sum = findrs.RecordCount
 If sum = 0 Then
    Exit Sub
 End If
Dim my()
ReDim my(1 To sum, 1 To 3)
findrs.MoveFirst
For i = 1 To sum
 my(i, 1) = "设备号" & findrs("machinenumber")  'labels
 my(i, 2) = findrs("occupy") 'series1 values
' my(i, 3) = findrs("totalamount")
 findrs.MoveNext
Next
  findrs.MoveFirst
 MSChart2.ChartData = my
 MSChart2.TitleText = cmbmachine1.Text & "设备负荷图示"
 MSChart2.Legend.VtFont.size = 14
 MSChart2.Title.VtFont.size = 14

End Sub
Private Sub Command1_Click()
  If rs.State = 1 Then Set rs = Nothing
  Unload Me
End Sub
 Private Sub Comok_Click()
    Dim rs1 As New ADODB.Recordset  '用rs来对零件计划标排序并显示之
    'machine()代表机器编号，machine2()代表机器效率machine1()代表机器负荷
    Dim machine() As String, machine1() As Single, machine2() As Single
    Dim timeoccupy  '记录每台设备的时间占用
    Dim rs2 As New ADODB.Recordset '找出相应的产品对应的加工设备号
    Dim rs3 As New ADODB.Recordset '打开打开日生产计划表
    Dim rs5 As New ADODB.Recordset '用来给r表里添加数据
    Dim m As Integer, i As Integer
    Dim quota As Single
    Dim kkk As Integer
    Dim Strg As String
    Dim conn As New ADODB.Connection
If Comok.Caption = "任务分派" Then
    Strg = Cmbsuanfa.Text
    Select Case Strg
      Case "算法1"
        kkk = genetic1
      Case "算法2"
        kkk = GA_TS
     
    End Select

    Set rs = Nothing
  rs.ActiveConnection = "dsn=dbw;uid=sa"
  rs.CursorLocation = adUseClient
  rs.CursorType = adOpenKeyset
  rs.LockType = adLockOptimistic
  rs.Source = "t_submachineload"
      conn = "dsn=dbw;uid=sa"
    conn.Open
    conn.Execute "DELETE  FROM t_submachineload "

  rs.Open
  
  Set DataGrid1.DataSource = rs
  Call first(DataGrid1)
  Call initial(DataGrid1, "设备负荷表1")
    Set rs1 = Nothing
    rs1.ActiveConnection = "dsn=dbw;uid=sa"
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenKeyset
    rs1.LockType = adLockOptimistic
    rs1.Source = "select workcode,ordercode, drawingno,processno,planquantity,machinenumber,timeoccupy " & _
    "from t_myplantask ,t_machineprocess1 where t_machineprocess1.drawingnumber=t_myplantask.drawingno" & _
    " and t_machineprocess1.processnumber=t_myplantask.processno  and added='否' order by drawingno,processno"
    rs1.Open
    rs1.MoveFirst
  For i = 1 To mm
      j = 1
    
      
      Do Until j > hh Or mach1(kkk, i, j).gx = 0
      rs1.Move CLng(mach1(kkk, i, j).gx - 1), adBookmarkFirst
      
      rs.AddNew
      rs("workcode") = CStr(rs1("workcode"))
      rs("ordercode") = CStr(rs1("ordercode"))
      rs("drawingnumber") = CStr(rs1("drawingno"))
      rs("processnumber") = CInt(rs1("processno"))
      rs("machinenumber") = CStr(rs1("machinenumber"))
      rs("quantity") = CInt(rs1("planquantity"))
      rs("timeoccupy") = CInt(rs1("planquantity")) * CSng(rs1("timeoccupy"))
      rs("plandate") = mach1(kkk, i, j).start
      rs("note") = mach1(kkk, i, j).stop
      'rs.MoveNext
      rs.Update
     ' rs.Close
      
      j = j + 1
      Loop
  Next i

Else
    frmsfcsh.Show
End If
        
End Sub

Private Sub Form_Load()
combb.Enabled = False
 Set rs = Nothing
  rs.ActiveConnection = "dsn=dbw;uid=sa"
  rs.CursorLocation = adUseClient
  rs.CursorType = adOpenKeyset
  rs.LockType = adLockOptimistic
  rs.Source = "t_submachineload"
  rs.Open
  Set DataGrid1.DataSource = rs
  Call first(DataGrid1)
  Call initial(DataGrid1, "设备负荷表")
  SSTab1.Tab = 0
  rs.Close
End Sub

Private Sub Label3_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim findrs As New ADODB.Recordset
  Dim i As Integer
If SSTab1.Tab = 0 Then
   
      Set rs = Nothing
        rs.ActiveConnection = "dsn=dbw;uid=sa"
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenKeyset
        rs.LockType = adLockOptimistic
        rs.Source = "t_submachineload"
        rs.Open
        Set DataGrid1.DataSource = rs
        Call first(DataGrid1)
        Call initial(DataGrid1, "设备负荷表")
  ElseIf SSTab1.Tab = 1 Then
       Set rs = Nothing
       rs.ActiveConnection = "dsn=dbw;uid=sa"
       rs.CursorLocation = adUseClient
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "select distinct machinenumber,devicename  from device ,t_submachineload  where machinenumber=deviceno order by machinenumber"
       rs.Open
       If rs.RecordCount <> 0 Then
         cmbmachine.Clear
         rs.MoveFirst
         Do Until rs.EOF
            cmbmachine.AddItem rs("machinenumber") & "/" & rs("devicename")
            rs.MoveNext
         Loop
        End If
Else
  If str1 <> "" Then
       Set rs = Nothing
       rs.ActiveConnection = "dsn=dbw;uid=sa"
       rs.CursorLocation = adUseClient
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "select distinct  devicename,machinenumber  from device,t_submachineload where machinenumber=deviceno order by devicename"
       rs.Open
       If rs.RecordCount <> 0 Then
         ReDim tt(1 To rs.RecordCount, 1 To 3)
         rs.MoveFirst
         i = 1
         Do Until rs.EOF
          Set findrs = Nothing
          findrs.ActiveConnection = "dsn=dbw;uid=sa"
          findrs.CursorLocation = adUseClient
          findrs.CursorType = adOpenKeyset
          findrs.LockType = adLockOptimistic
          findrs.Source = "select top 1 devicename,t_submachineload.note  from" & _
          " t_submachineload,device where  machinenumber=deviceno and machinenumber='" & Trim$(rs("machinenumber")) & "' order by t_submachineload.note desc"
          findrs.Open
          If findrs.RecordCount <> 0 Then
           tt(i, 1) = CStr(findrs("devicename"))
          If findrs("note") <> Null Then
           tt(i, 2) = CSng(findrs("note"))
           Else
           tt(i, 2) = 5
           End If
           i = i + 1
          End If
          rs.MoveNext
        Loop
        MSChart2.ChartData = tt
        MSChart2.chartType = VtChChartType2dBar
      End If
  Else
       Set rs = Nothing
       rs.ActiveConnection = "dsn=dbw;uid=sa"
       rs.CursorLocation = adUseClient
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "select distinct  devicename  from device,t_submachineload where machinenumber=deviceno order by devicename"
       rs.Open
       If rs.RecordCount <> 0 Then
         cmbmachine1.Clear
         rs.MoveFirst
         Do Until rs.EOF
            cmbmachine1.AddItem rs("devicename")
            rs.MoveNext
         Loop
        End If
  End If
End If
End Sub




