VERSION 5.00
Begin VB.Form Regist 
   Caption         =   "Registrasi"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   Picture         =   "Regist.frx":0000
   ScaleHeight     =   5715
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Pesan Tiket Pesawat"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Booking Kamar Hotel"
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3720
      TabIndex        =   2
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   3720
      TabIndex        =   1
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3720
      TabIndex        =   0
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registrasi"
      BeginProperty Font 
         Name            =   "OCR A Std"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   1680
      TabIndex        =   6
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "No. KTP  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Tlp    :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama      :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5760
      Left            =   0
      Picture         =   "Regist.frx":42762
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9930
   End
End
Attribute VB_Name = "Regist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Database2.Text1 = Text1.Text
Database2.Text2 = Text2.Text
Database2.Text3 = Text3.Text
Output2.Label15 = Text1.Text
Output2.Label27 = Text1.Text
Output2.Label28 = Text2.Text
Output2.Label29 = Text3.Text
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox ("Data Harus Lengkap"), vbInformation, "WARNING!"
Else
Hotel.Show
Unload Me
End If
End Sub

Private Sub Command2_Click()
Database.Text1 = Text1.Text
Database.Text2 = Text2.Text
Database.Text3 = Text3.Text
Output.Label15 = Text1.Text
Output.Label27 = Text1.Text
Output.Label28 = Text2.Text
Output.Label29 = Text3.Text
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox ("Data Harus Lengkap"), vbInformation, "WARNING!"
Else
Form1.Show
Unload Me
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub
