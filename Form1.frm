VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tipe Perjalanan"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Pulang - Pergi"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sekali Berangkat"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Tipe Perjalanan Anda"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   5400
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Tkt_Pswt.Show
Unload Me
End Sub

Private Sub Command2_Click()
Tkt_Pswt2.Show
Unload Me
End Sub

