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
        '��¼ʧ�ܣ��˳�Ӧ�ó���
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
  Case "������"
        dg.Columns(0).Caption = "������"
        dg.Columns(1).Caption = "������"
        dg.Columns(2).Caption = "�����ͺ�"
        dg.Columns(3).Caption = "��Ʒ����"
        dg.Columns(4).Caption = "ͼ��"
        dg.Columns(5).Caption = "����"
        dg.Columns(6).Caption = "Ԥ������"
        dg.Columns(7).Caption = "��������"
        dg.Columns(8).Caption = "����ƻ���"
        dg.Columns(9).Caption = "��ע"
   Case "֪ʶ��"
        dg.Columns(0).Caption = "�����"
        dg.Columns(1).Caption = "��������"
        dg.Columns(2).Caption = "��������"
        dg.Columns(3).Caption = "����ģ��"
        dg.Columns(4).Caption = "����Ŀ��"
        dg.Columns(5).Caption = "�㷨����"
  Case "�㷨��"
        dg.Columns(0).Caption = "�㷨����"
        dg.Columns(1).Caption = "�㷨����"
        dg.Columns(2).Caption = "����ģ��"
        dg.Columns(3).Caption = "�㷨����"
        dg.Columns(4).Caption = "�㷨����"
        dg.Columns(5).Caption = "��ע"
        
  Case "�����"
        dg.Columns(0).Caption = "������"
        dg.Columns(1).Caption = "�����ͺ�"
        dg.Columns(2).Caption = "������"
        dg.Columns(3).Caption = "����ͼ��"
        dg.Columns(4).Caption = "ͼ��"
        dg.Columns(5).Caption = "̨������"
        dg.Columns(6).Caption = "�ƻ�����"
        dg.Columns(7).Caption = "��������"
        dg.Columns(8).Caption = "��������"
        dg.Columns(9).Caption = "����ƻ���"
        dg.Columns(10).Caption = "��ע"
    Case "����������ƻ�"
        dg.Columns(0).Caption = "������"
        dg.Columns(1).Caption = "������"
        dg.Columns(2).Caption = "����ͼ��"
        dg.Columns(3).Caption = "ͼ��"
        dg.Columns(4).Caption = "���պ�"
        dg.Columns(5).Caption = "̨������"
        dg.Columns(6).Caption = "�ƻ�����"
        dg.Columns(7).Caption = "��������"
        dg.Columns(8).Caption = "��������"
        dg.Columns(9).Caption = "����ƻ���"
        dg.Columns(10).Caption = "��ע"

  Case "�������ƻ�"
         dg.Columns(0).Caption = "������"
         dg.Columns(1).Caption = "������"
         dg.Columns(2).Caption = "��Ʒͼ��"
         dg.Columns(3).Caption = "�����"
         dg.Columns(4).Caption = "������"
         dg.Columns(5).Caption = "�豸��"
         dg.Columns(6).Caption = "��ע"
  Case "�豸���ɱ�"
         dg.Columns(0).Caption = "������"
         dg.Columns(1).Caption = "������"
         dg.Columns(2).Caption = "�豸��"
         dg.Columns(3).Caption = "��Ʒͼ��"
         dg.Columns(4).Caption = "������"
         dg.Columns(5).Caption = "�����"
         dg.Columns(6).Caption = "ռ��ʱ��"
         dg.Columns(7).Caption = "�ƻ�����"
         dg.Columns(8).Caption = "��ע"
Case "�豸���ɱ�1"
         dg.Columns(0).Caption = "������"
         dg.Columns(1).Caption = "������"
         dg.Columns(2).Caption = "�豸��"
         dg.Columns(3).Caption = "��Ʒͼ��"
         dg.Columns(4).Caption = "������"
         dg.Columns(5).Caption = "�����"
         dg.Columns(6).Caption = "ռ��ʱ��"
         dg.Columns(7).Caption = "����ʱ��"
         dg.Columns(8).Caption = "����ʱ��"
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
    Dim bond As Integer, cnt(100) As Integer                 'bondΪ��̬������Ͻ����
    Dim i As Integer, cno As Integer    'cnoΪ�жϲ�Ʒ��̨��,����·�������ת������

    tv.Nodes.Clear
    connbom = "DSN=dlrwdb;uid=scl;uid=scl"
    connbom.Open
    tempnode = tv.Nodes.Add(, , "��ƷBOMͼ", "��ƷBOMͼ" & " ")
    cno = snumber
    '�Ѳ�����ֵ��opennode����
    ReDim Preserve opennode(1)
    With opennode(1)
        .father = sdrawingno
        .name = sname
        .drawingno = sdrawingno
        .number = snumber / snumber
    End With
    
    '��Ӳ�������ʾ�ĵ�һ���ڵ�
    bond = 1
    stemp = sdrawingno & "(" & sname & ")" & " " & snumber / cno
    nodekey = Trim(sdrawingno) & Trim(sdrawingno)
     Set tempnode = tv.Nodes.Add("��ƷBOMͼ", tvwChild, nodekey, stemp)
     tempnode.EnsureVisible
    '����ӽڵ�,�����ֽڵ����opennode������
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
         '�˴��ж��Ƿ���ͼ������ֱϵ����ͼ���ظ�
         If layer <> 1 Then '��һ�㲻���ж�
            For k = 1 To layer
             If prodrawno(k) = opennode(bond).drawingno Then
               warnstr = "��ͼ�� " & opennode(bond).name & " | " & opennode(bond).drawingno & " ����� " & k & "������ͼ�� " & prodrawno(k)
               warnstr = warnstr & " �ظ����������ѭ���������˳�"
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
   'Ϊ������ֵ,���رռ�¼
   showbom = opennode
  MsgBox "�� " & i & "����¼", vbOKOnly
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

