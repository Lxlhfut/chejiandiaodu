VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "欢迎进入车间生产调度系统"
   ClientHeight    =   7815
   ClientLeft      =   75
   ClientTop       =   1710
   ClientWidth     =   11160
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3120
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16869
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16BBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1700D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1735F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":177B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17C03
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17F55
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":183A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":187F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18C4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1909D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":193EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19741
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19B93
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19EAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A1C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A321
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A773
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AA8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ADA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B1F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B64B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BA9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BDB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C209
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7545
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14023
            Text            =   "状态"
            TextSave        =   "状态"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2008-4-24"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "20:32"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_order 
      Caption         =   "作业计划"
   End
   Begin VB.Menu mnu_chj 
      Caption         =   "作业拆解"
   End
   Begin VB.Menu mnuTaskDes 
      Caption         =   "任务倒排"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuTaskSort 
      Caption         =   "任务排序及分配"
      Begin VB.Menu mnu_zdfp 
         Caption         =   "自动分配"
      End
   End
   Begin VB.Menu mnuSee 
      Caption         =   "查看"
      Visible         =   0   'False
      Begin VB.Menu mnuSeeBigIcon 
         Caption         =   "大图标"
      End
      Begin VB.Menu mnuSeeSmallIcon 
         Caption         =   "小图标"
      End
      Begin VB.Menu mnuSeeDetail 
         Caption         =   "详细资料"
      End
      Begin VB.Menu mnuSeeList 
         Caption         =   "列表"
      End
   End
   Begin VB.Menu mnuHelpAbout 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)



Private Sub frmctrlbom_Click()
frmbomctrl.Show
End Sub






Private Sub GA_TS11_Click()

End Sub

Private Sub genetic_example_Click()
实例2.Show
End Sub

Private Sub genetic_menu_Click()
frmsuanfa1.Show
End Sub







Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub


Private Sub menusort_Click()
  frmgeneralpartplan1.Show
End Sub

Private Sub mnu_chj_Click()
   frmgeneralpartplan1.Show
End Sub



Private Sub mnu_order_Click()
   Frmorder1.Show
End Sub

Private Sub mnu_shgfp_Click()
    Dim i As Integer
i = getPower("2")
If i = 1 Then
    frmPower.Show vbModal
    If frmPower.OK_power Then
    i = 0
    Unload frmPower
    End If
End If
If i = 0 Then
    fMainForm.mnuTaskSort.Enabled = False
    frmTaskdist.Show
End If
End Sub

Private Sub mnu_zdfp_Click()
      frmsuanfa1.Show
End Sub





Private Sub mnudevice_Click()
frmCtrlDevice.Show
End Sub

Private Sub mnudeviceclass_Click()
frmCtrldeviceclass.Show
End Sub

Private Sub mnuSeeBigIcon_Click()
frmWorkshop.LVDevice.View = lvwIcon
frmWorkshop.LVDeviceClass.View = lvwIcon
mnuSeeBigIcon.Checked = True
mnuSeeDetail.Checked = False
mnuSeeList.Checked = False
mnuSeeSmallIcon.Checked = False
End Sub

Private Sub mnuSeeDetail_Click()
frmWorkshop.LVDevice.View = lvwReport
frmWorkshop.LVDeviceClass.View = lvwReport
mnuSeeDetail.Checked = True
mnuSeeList.Checked = False
mnuSeeSmallIcon.Checked = False
mnuSeeBigIcon.Checked = False
End Sub

Private Sub mnuSeeList_Click()
frmWorkshop.LVDevice.View = lvwList
frmWorkshop.LVDeviceClass.View = lvwList
mnuSeeList.Checked = True
mnuSeeSmallIcon.Checked = False
mnuSeeBigIcon.Checked = False
mnuSeeDetail.Checked = False
End Sub

Private Sub mnuSeeSmallIcon_Click()
frmWorkshop.LVDevice.View = lvwSmallIcon
frmWorkshop.LVDeviceClass.View = lvwSmallIcon
mnuSeeSmallIcon.Checked = True
mnuSeeList.Checked = False
mnuSeeBigIcon.Checked = False
mnuSeeDetail.Checked = False
End Sub

Private Sub mnusfgz_Click()
  Form1.Show
End Sub




Private Sub mnuSpect_Click()
fMainForm.mnuSpect.Enabled = False
frmSpect.Show
End Sub

Private Sub mnuSysShop1_Click()
frmworkshop1.Show
End Sub

Private Sub mnuSysWorkshop_Click()
'fMainForm.mnuSysWorkshop.Enabled = False
'frmMain.mnuSysWorkshop.Enabled = False
'frmWorkshop.Show
frmCtrlworkshop.Show
End Sub

Private Sub mnuTaskDes_Click()
fMainForm.mnuTaskDes.Enabled = False
frmTaskDes.Show
End Sub

Private Sub mnuTaskSend_Click()
'fMainForm.mnuTaskSend.Enabled = False
'frmMain.mnuTaskSend.Enabled = False
frmTaskSend.Show
End Sub



Private Sub mnuTest_Click()
Frmtest.Show
End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "sysadm"
            fMainForm.mnuSystemAdm.Enabled = False
            frmSystemAdm.Show
        Case "workshop"
            fMainForm.mnusysworkshop.Enabled = False
            frmWorkshop.Show
        Case "changeGroup"
            'fMainForm.mnuSysWorkshop.Enabled = False
            'frmWorkshop.Show
        Case "send"
            fMainForm.mnuTaskSend.Enabled = False
            frmTaskSend.Show
        Case "dist"
            fMainForm.mnuTaskSort.Enabled = False
            frmTaskdist.Show
        Case "spect"
            fMainForm.mnuSpect.Enabled = False
            frmSpect.Show
        Case "dataadm"
            fMainForm.mnuData.Enabled = False
            frmDataAdm.Show
            
        Case "quit"
            Call mnuSystemExit_Click
        Case "test"
            'fMainForm.mnuTest = False
           ' Frmtest.Show
            
            'mnuFilePrint_Click
        Case "algrithm"
            frmsuanfa1.Show
        Case "复制"
            'mnuEditCopy_Click
        Case "粘贴"
            'mnuEditPaste_Click
        Case "粗体"
            'ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
            'Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
        Case "斜体"
            'ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
            'Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
        Case "下划线"
            'ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
            'Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "左对齐"
            'ActiveForm.rtfText.SelAlignment = rtfLeft
        Case "置中"
            'ActiveForm.rtfText.SelAlignment = rtfCenter
        Case "右对齐"
            'ActiveForm.rtfText.SelAlignment = rtfRight
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub






Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub



Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub


Private Sub tbToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
    Case "buff"
        FrmAlgBuff.Tag = "buff"
        FrmAlgBuff.Show
    Case "prodtime"
        FrmAlgBuff.Tag = "prodtime"
        FrmAlgBuff.Show
    Case "ljb"
        FrmAlgBuff.Tag = "ljb"
        FrmAlgBuff.Show
    Case "other"
        FrmAlgBuff.Tag = "other"
        FrmAlgBuff.Show
    Case "ddjh"
        Frmorder1.Show
    Case "jhcj"
        frmgeneralpartplan1.Show
End Select
End Sub

Function getPower(pow As String) As Integer
Dim sql As String, rs As New ADODB.Recordset
Dim mconn As New ADODB.Connection

Dim i As Integer
mconn.Open "DSN=dlrwdb;uid=scl"
sql = "select * from passwd where username='" & CurrentUser & "'"
rs.Open sql, mconn, adOpenKeyset, adLockPessimistic
If rs.RecordCount = 0 Then
    MsgBox "系统异常，联系系统管理员", vbOKOnly
    rs.Close
    getPower = 1
    Exit Function
End If
If InStr(1, rs("power"), pow, vbTextCompare) Then
    rs.Close
    getPower = 0
    Exit Function
End If
getPower = 1
Set mconn = Nothing
End Function



