VERSION 5.00
Begin VB.Form Main_Menu 
   Caption         =   "Traveloman"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   9360
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   5895
      Left            =   0
      Picture         =   "Main_Menu.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "Main_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Regist.Show
Unload Me
End Sub
