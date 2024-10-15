VERSION 5.00
Begin VB.Form frmsfcsh 
   Caption         =   "参数设置"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   ScaleHeight     =   3915
   ScaleWidth      =   5550
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      Begin VB.TextBox txtss 
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
         Left            =   1680
         TabIndex        =   5
         Text            =   "40"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtpc 
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
         Left            =   1680
         TabIndex        =   4
         Text            =   "0.89"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtpm 
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
         Left            =   1680
         TabIndex        =   3
         Text            =   "0.35"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtdd 
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
         Left            =   1680
         TabIndex        =   2
         Text            =   "60"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Comok 
         Caption         =   "确定"
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "种群大小"
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
         Left            =   720
         TabIndex        =   9
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复制率"
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
         Left            =   930
         TabIndex        =   8
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "变异率"
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
         Left            =   930
         TabIndex        =   7
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "迭代代数"
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
         Left            =   720
         TabIndex        =   6
         Top             =   2040
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmsfcsh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Comok_Click()
  If txtss.Text = "" Then
    MsgBox "请你输入种群大小", vbExclamation + vbInformation
    Exit Sub
  End If
  If txtpc.Text = "" Then
    MsgBox "请你输入复制率", vbExclamation + vbInformation
    Exit Sub
  End If
   If txtpm.Text = "" Then
    MsgBox "请你输入变异率", vbExclamation + vbInformation
    Exit Sub
  End If
   If txtdd.Text = "" Then
    MsgBox "请你输入迭代代数", vbExclamation + vbInformation
    Exit Sub
  End If
   ss1 = CInt(txtss.Text)
   pc1 = CSng(txtpc.Text)
   pm1 = CSng(txtpm.Text)
   dd1 = CInt(txtdd.Text) '
   frmsuanfa1.Comok.Caption = "任务分派"
   Unload Me
End Sub

