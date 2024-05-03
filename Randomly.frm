VERSION 5.00
Begin VB.Form Randomly 
   Caption         =   "随机抽号"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5550
   LinkTopic       =   "Form3"
   ScaleHeight     =   3615
   ScaleWidth      =   5550
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Text            =   "53"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   720
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "人数："
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Randomly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "开始" Then
    Command1.Caption = "停止"
    Timer1.Enabled = True
Else
    Command1.Caption = "开始"
Timer1.Enabled = False
    List1.AddItem Label1.Caption
End If
End Sub

Private Sub Command2_Click()
Form1.Show
Randomly.Hide
End Sub

Private Sub Timer1_Timer()
X = Int(Rnd * 53) + 1
Label1.Caption = X
End Sub
