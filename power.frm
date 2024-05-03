VERSION 5.00
Begin VB.Form power 
   Caption         =   "电源操作"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4035
   LinkTopic       =   "Form3"
   ScaleHeight     =   1845
   ScaleWidth      =   4035
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "关机"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重启"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "想立即关机还是计划重启？？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "操作前请保存好您未保存的工作。"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "power"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "cmd.exe /c shutdown -s -t 0"
End Sub

Private Sub Command2_Click()
Shell "cmd.exe /c shutdown -a -t 0"
End Sub
