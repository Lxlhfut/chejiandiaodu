Attribute VB_Name = "Module1"
Option Explicit
Type mytype
    father As String
    name As String
    type As String
    number As Integer
    renewid As String
    drawingno As String
End Type
Public ss1 As Integer
Public pc1 As Single
Public pm1 As Single
Public dd1 As Integer
Public fMainForm As frmMain
Public gConnstr As String, gConnstr_scl As String
Public gIconAmount As Integer
Public Type ProcessMatrix
     processname As String
     deviceno As String
     otime As Single
End Type
Public Type Vertex
    'cost As Single
    vtxno As String
    vtxname As String
End Type

Public CurrentUser As String, Period As Integer
Option Base 0
Public Type machine
       gx As Integer
       start As Single
       stop As Single
End Type


Sub Main()
   Dim fLogin As New frmLogin
'    gConnstr = "dlrwdb"
    gConnstr = "dbw"
    gIconAmount = 4
    fLogin.Show vbModal
    If Not fLogin.OK Then
        '登录失败，退出应用程序
        End
    End If
    Unload fLogin


    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    'Load fMainForm
    Unload frmSplash


    fMainForm.Show
    'frmWorkshop.Show
    'Form1.Show
End Sub
Public Sub initial(dg As DataGrid, table As String)
  Select Case table
  Case "定单表"
        dg.Columns(0).Caption = "定单号"
        dg.Columns(1).Caption = "工作号"
        dg.Columns(2).Caption = "机车型号"
        dg.Columns(3).Caption = "产品名称"
        dg.Columns(4).Caption = "图号"
        dg.Columns(5).Caption = "数量"
        dg.Columns(6).Caption = "预收日期"
        dg.Columns(7).Caption = "交货日期"
        dg.Columns(8).Caption = "加入计划否"
        dg.Columns(9).Caption = "备注"
   Case "知识库"
        dg.Columns(0).Caption = "规则号"
        dg.Columns(1).Caption = "规则名称"
        dg.Columns(2).Caption = "问题描述"
        dg.Columns(3).Caption = "车间模型"
        dg.Columns(4).Caption = "调度目标"
        dg.Columns(5).Caption = "算法代号"
  Case "算法库"
        dg.Columns(0).Caption = "算法代号"
        dg.Columns(1).Caption = "算法名称"
        dg.Columns(2).Caption = "车间模型"
        dg.Columns(3).Caption = "算法参数"
        dg.Columns(4).Caption = "算法描述"
        dg.Columns(5).Caption = "备注"
        
  Case "零件表"
        dg.Columns(0).Caption = "工作号"
        dg.Columns(1).Caption = "机车型号"
        dg.Columns(2).Caption = "定单号"
        dg.Columns(3).Caption = "父件图号"
        dg.Columns(4).Caption = "图号"
        dg.Columns(5).Caption = "台分数量"
        dg.Columns(6).Caption = "计划数量"
        dg.Columns(7).Caption = "订货日期"
        dg.Columns(8).Caption = "交货日期"
        dg.Columns(9).Caption = "加入计划否"
        dg.Columns(10).Caption = "备注"
    Case "工艺与零件计划"
        dg.Columns(0).Caption = "工作号"
        dg.Columns(1).Caption = "定单号"
        dg.Columns(2).Caption = "父件图号"
        dg.Columns(3).Caption = "图号"
        dg.Columns(4).Caption = "工艺号"
        dg.Columns(5).Caption = "台分数量"
        dg.Columns(6).Caption = "计划数量"
        dg.Columns(7).Caption = "订货日期"
        dg.Columns(8).Caption = "交货日期"
        dg.Columns(9).Caption = "加入计划否"
        dg.Columns(10).Caption = "备注"

  Case "日生产计划"
         dg.Columns(0).Caption = "工作号"
         dg.Columns(1).Caption = "订单号"
         dg.Columns(2).Caption = "产品图号"
         dg.Columns(3).Caption = "工序号"
         dg.Columns(4).Caption = "工作量"
         dg.Columns(5).Caption = "设备号"
         dg.Columns(6).Caption = "备注"
  Case "设备负荷表"
         dg.Columns(0).Caption = "工作号"
         dg.Columns(1).Caption = "订单号"
         dg.Columns(2).Caption = "设备号"
         dg.Columns(3).Caption = "产品图号"
         dg.Columns(4).Caption = "任务量"
         dg.Columns(5).Caption = "工序号"
         dg.Columns(6).Caption = "占用时间"
         dg.Columns(7).Caption = "计划日期"
         dg.Columns(8).Caption = "备注"
Case "设备负荷表1"
         dg.Columns(0).Caption = "工作号"
         dg.Columns(1).Caption = "订单号"
         dg.Columns(2).Caption = "设备号"
         dg.Columns(3).Caption = "产品图号"
         dg.Columns(4).Caption = "任务量"
         dg.Columns(5).Caption = "工序号"
         dg.Columns(6).Caption = "占用时间"
         dg.Columns(7).Caption = "开工时间"
         dg.Columns(8).Caption = "结束时间"
  End Select
  
End Sub

Public Sub first(dg As DataGrid)
 Dim col As Integer
 Dim wide
 Dim i As Integer
 col = dg.Columns.Count
 wide = dg.Width
 wide = wide / col
 For i = 0 To col - 1
    dg.Columns(i).Width = wide
 Next
 End Sub
Public Sub init()
 Dim tt As Control
  For Each tt In Me.Controls
     If TypeOf tt Is TextBox Then
       tt = ""
     ElseIf TypeOf tt Is ComboBox Then
       tt.Text = ""
     End If
  Next
End Sub

Public Function showbom(sdrawingno As String, sname As String, _
tv As TreeView, snumber As Integer) As mytype()
    On Error Resume Next
    Dim tempnode As Node
    Dim stemp As String, warnstr As String
    Dim k As Integer
    Dim connbom As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim nodekey As String
    Dim f, m
    Dim opennode() As mytype
    Dim bond As Integer, cnt(100) As Integer                 'bond为动态数组的上界变量
    Dim i As Integer, cno As Integer    'cno为判断产品是台份,还是路用配件的转换参数

    tv.Nodes.Clear
    connbom = "DSN=dlrwdb;uid=scl;uid=scl"
    connbom.Open
    tempnode = tv.Nodes.Add(, , "产品BOM图", "产品BOM图" & " ")
    cno = snumber
    '把参数赋值给opennode数组
    ReDim Preserve opennode(1)
    With opennode(1)
        .father = sdrawingno
        .name = sname
        .drawingno = sdrawingno
        .number = snumber / snumber
    End With
    
    '添加参数所表示的第一个节点
    bond = 1
    stemp = sdrawingno & "(" & sname & ")" & " " & snumber / cno
    nodekey = Trim(sdrawingno) & Trim(sdrawingno)
     Set tempnode = tv.Nodes.Add("产品BOM图", tvwChild, nodekey, stemp)
     tempnode.EnsureVisible
    '添加子节点,并把字节点加入opennode数组中
    Dim strsql As String
    Dim tmptype
    Dim tmpdrawingno As String, layer As Integer
    Dim prodrawno() As String, pt As String, chd As String
    Dim flag As Boolean
    'tmptype = opennode(bond).type
    tmpdrawingno = sdrawingno
    i = 1
    layer = 1
    cnt(1) = 1
    f = layer
    flag = False
    ReDim prodrawno(1)
    prodrawno(1) = Trim(sdrawingno)
    Do Until bond = 0
        While cnt(layer) = 0
            layer = layer - 1
        Wend
        pt = ""
        For m = 1 To layer
            pt = pt & prodrawno(m)
        Next m
        stemp = opennode(bond).drawingno & "(" & opennode(bond).name & ")" & " " & opennode(bond).number
        chd = pt & opennode(bond).drawingno
        
        Set tempnode = tv.Nodes.Add(pt, tvwChild)
        tempnode.Key = chd
        tempnode.Text = stemp
        'tempnode.EnsureVisible
         '此处判断是否子图号与其直系祖先图号重复
         If layer <> 1 Then '第一层不做判断
            For k = 1 To layer
             If prodrawno(k) = opennode(bond).drawingno Then
               warnstr = "子图号 " & opennode(bond).name & " | " & opennode(bond).drawingno & " 与其第 " & k & "层祖先图号 " & prodrawno(k)
               warnstr = warnstr & " 重复，会造成死循环，程序退出"
               MsgBox warnstr, vbOKOnly
               rs.Close
               bond = 1
               k = layer + 1
             End If
            Next k
        End If
        cnt(layer) = cnt(layer) - 1
        tmptype = opennode(bond).type
        tmpdrawingno = opennode(bond).drawingno
        strsql = "select * from t_bom where father='" & tmpdrawingno & "'" & " and son<>'empty' "
        rs.Open strsql, connbom, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            flag = True
            layer = layer + 1
            cnt(layer) = rs.RecordCount
            ReDim Preserve prodrawno(UBound(prodrawno) + 1)
            prodrawno(layer) = tmpdrawingno
            rs.MoveFirst
            While Not rs.EOF
                 ReDim Preserve opennode(bond)
                 With opennode(bond)
                          .father = Trim(rs("father"))
                          .name = Trim(rs("sname"))
                          .drawingno = Trim(rs("son"))
                          .number = rs("pargroupamount") / cno
                          If (.number = 0) Then
                            .number = rs("pargroupamount")
                          End If
                          
                End With
                bond = bond + 1
                i = i + 1
                rs.MoveNext
            Wend
        End If
        rs.Close
        bond = bond - 1
   Loop
   '为函数赋值,并关闭记录
   showbom = opennode
  MsgBox "共 " & i & "条记录", vbOKOnly
   connbom.Close
   Set rs = Nothing
   Set connbom = Nothing
End Function

Sub setpara(pa As String)
Dim sql As String, rs As New ADODB.Recordset, mconn As New ADODB.Connection
mconn.Open "DSN=dlrwdb;uid=scl;uid=scl"
sql = "select * from t_ctrl"
rs.CursorLocation = adUseClient
rs.Open sql, mconn, adOpenKeyset, adLockPessimistic
rs.MoveFirst
rs("note2") = pa
rs.Update
rs.Close

End Sub

Function getpara() As String
Dim sql As String, rs As New ADODB.Recordset, mconn As New ADODB.Connection
mconn.Open "DSN=dlrwdb;uid=sa"
sql = "select * from t_ctrl"
rs.CursorLocation = adUseClient
rs.Open sql, mconn, adOpenKeyset, adLockPessimistic
rs.MoveFirst
getpara = rs("note2")
rs.Close
End Function

