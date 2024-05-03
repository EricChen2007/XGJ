VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tools"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   3765
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "功能区"
      Height          =   2535
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   1575
      Begin VB.CommandButton Command2 
         Caption         =   "抽人"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "敬请期待。。。"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "快捷指令"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      Begin VB.CommandButton Command4 
         Caption         =   "电源操作"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "PowerShell"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CMD"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "cmd.exe cmd"
End Sub

Private Sub Command2_Click()
Randomly.Show
End Sub

Private Sub Command3_Click()
Shell "powershell.exe powershell"
End Sub

Private Sub Command4_Click()
power.Show
End Sub

Private Sub Label1_Click()

End Sub
