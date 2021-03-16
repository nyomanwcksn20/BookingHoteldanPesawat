VERSION 5.00
Begin VB.Form Tkt_Pswt 
   Caption         =   "Tiket Pesawat"
   ClientHeight    =   6855
   ClientLeft      =   8160
   ClientTop       =   735
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   Picture         =   "Tiket_Pesawat.frx":0000
   ScaleHeight     =   6855
   ScaleWidth      =   11580
   Begin VB.ListBox List4 
      Height          =   450
      Left            =   4920
      TabIndex        =   29
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   28
      Top             =   5280
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Selanjutnya"
      Height          =   495
      Left            =   9960
      TabIndex        =   24
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Batal"
      Height          =   495
      Left            =   6720
      TabIndex        =   23
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   20
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   19
      Top             =   3840
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4680
      Width           =   3135
   End
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   3840
      TabIndex        =   17
      Top             =   3120
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   2520
      TabIndex        =   16
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1560
      TabIndex        =   15
      Top             =   3120
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080FF80&
      Caption         =   "Bisnis"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   2400
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FF80&
      Caption         =   "Ekonomi"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   2400
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Text            =   "Combo3"
      Top             =   1680
      Width           =   4215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Text            =   "Combo2"
      Top             =   960
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   240
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Bagasi"
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   5880
      Width           =   5655
      Begin VB.OptionButton Option5 
         BackColor       =   &H0080FF80&
         Caption         =   "5 Kg >"
         Height          =   495
         Left            =   3840
         TabIndex        =   26
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080FF80&
         Caption         =   "0 - 5 Kg"
         Height          =   495
         Left            =   1080
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "120000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   33
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "70000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   32
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Potongan 70%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Potongan 50%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Anak - Anak"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4560
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   3735
      Left            =   6720
      Top             =   720
      Width           =   4575
   End
   Begin VB.Image Picture1 
      Height          =   3735
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   720
      Width           =   4575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   6240
      X2              =   6240
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Lihat Gambar"
      Height          =   255
      Left            =   6720
      TabIndex        =   22
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Harga Rp"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayi"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Dewasa"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Berangkat / Jam (WIB)"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Tiket"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ke"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dari"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Maskapai"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Tkt_Pswt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "Garuda" Then
Picture1.Picture = LoadPicture(App.Path & "\Garuda0.jpg")
ElseIf Combo1.Text = "Citilink" Then
Picture1.Picture = LoadPicture(App.Path & "\Citilink0.jpg")
ElseIf Combo1.Text = "Lion Air" Then
Picture1.Picture = LoadPicture(App.Path & "\Lion0.jpg")
ElseIf Combo1.Text = "Air Asia" Then
Picture1.Picture = LoadPicture(App.Path & "\Asia0.jpg")
End If
End Sub

Private Sub Command3_Click()
Main_Menu.Show
Unload Me
End Sub

Private Sub Command4_Click()
Database.Text5 = Combo2.Text
Database.Text6 = Combo3.Text
Database.Text4 = Combo1.Text
Database.Text7 = List1.Text + List2.Text + List3.Text
Database.Text8 = "-"
Database.Text9 = Text2.Text
Database.Text10 = Text3.Text
Database.Text11 = Text4.Text
Database.Text12 = Text5.Text

Output.Label1 = Combo2.Text
Output.Label2 = Combo3.Text
Output.Label7 = Combo1.Text
Output.Label8 = List4.Text
Output.Label9 = List1.Text
Output.Label10 = List2.Text
Output.Label11 = List3.Text
Output.Label17 = Text3.Text
Output.Label20 = Text2.Text
Output.Label21 = Text4.Text
Output.Label30 = Text5.Text
If Option1.Enabled = True Then
Output.Label23 = "Ekonomi"
ElseIf Option2.Enabled = True Then
Output.Label23 = "Bisnis"
End If
Database.Show
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

For t = 2017 To 2050
List3.AddItem t
Next
For h = 1 To 31
List1.AddItem h
Next
List2.List(0) = "Januari"
List2.List(1) = "Februari"
List2.List(2) = "Maret"
List2.List(3) = "April"
List2.List(4) = "Mei"
List2.List(5) = "Juni"
List2.List(6) = "Juli"
List2.List(7) = "Agustus"
List2.List(8) = "September"
List2.List(9) = "Oktober"
List2.List(10) = "November"
List2.List(11) = "Desember"

List4.List(0) = "08:00"
List4.List(1) = "10:00"
List4.List(2) = "12:00"
List4.List(3) = "14:00"
List4.List(4) = "16:00"
List4.List(5) = "18:00"
List4.List(6) = "20:00"
List4.List(7) = "22:00"
List4.List(8) = "24:00"

Combo1.List(0) = "Garuda"
Combo1.List(1) = "Citilink"
Combo1.List(2) = "Lion Air"
Combo1.List(3) = "Air Asia"


Combo2.List(0) = "Jakarta"
Combo2.List(1) = "Surabaya"
Combo2.List(2) = "Bali"
Combo2.List(3) = "Aceh"

Combo3.List(0) = "Jakarta"
Combo3.List(1) = "Surabaya"
Combo3.List(2) = "Bali"
Combo3.List(3) = "Aceh"
End Sub


Private Sub Option1_Click()
If Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 1000000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 1200000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 1200000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 1300000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 800000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 850000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1100000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1500000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1700000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 1000000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1450000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1600000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 600000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 650000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 700000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 750000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 500000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 500000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1000000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1300000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1500000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 900000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1250000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1300000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 600000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 650000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 800000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 850000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 450000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 500000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 700000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1000000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1200000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 800000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 900000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1000000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 800000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 850000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 800000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 900000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 550000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 550000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 900000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1000000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1200000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 900000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1100000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1200000
ElseIf Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
MsgBox "Kota Tidak Boleh Sama", vbInformation, "WARNING!"
ElseIf Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
MsgBox "Kota Tidak Boleh Sama", vbInformation, "WARNING!"
ElseIf Combo2.ListIndex = 2 And Combo3.ListIndex = 2 Then
MsgBox "Kota Tidak Boleh Sama", vbInformation, "WARNING!"
ElseIf Combo2.ListIndex = 3 And Combo3.ListIndex = 3 Then
MsgBox "Kota Tidak Boleh Sama", vbInformation, "WARNING!"
End If

If Combo1.Text = "Garuda" Then
Picture1.Picture = LoadPicture(App.Path & "\Garuda1.jpg")
ElseIf Combo1.Text = "Citilink" Then
Picture1.Picture = LoadPicture(App.Path & "\Citilink1.jpg")
ElseIf Combo1.Text = "Lion Air" Then
Picture1.Picture = LoadPicture(App.Path & "\Lion1.jpg")
ElseIf Combo1.Text = "Air Asia" Then
Picture1.Picture = LoadPicture(App.Path & "\Asia1.jpg")
End If
End Sub

Private Sub Option2_Click()
If Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 1700000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 1650000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 2000000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 2300000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 1300000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 1400000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1500000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1700000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1800000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 1500000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1900000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 2000000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 1600000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 1500000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 1800000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 2000000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 900000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 1000000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1300000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1500000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1700000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 1150000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1500000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1800000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 800000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 900000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 1100000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 1150000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 700000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 800000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1000000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1100000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1500000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 1000000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1150000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1200000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 1150000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 1250000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 1100000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 1200000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 700000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 750000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1000000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1200000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1300000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 1050000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1500000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1600000
ElseIf Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
MsgBox "Kota Tidak Boleh Sama", vbInformation, "WARNING!"
ElseIf Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
MsgBox "Kota Tidak Boleh Sama", vbInformation, "WARNING!"
ElseIf Combo2.ListIndex = 2 And Combo3.ListIndex = 2 Then
MsgBox "Kota Tidak Boleh Sama", vbInformation, "WARNING!"
ElseIf Combo2.ListIndex = 3 And Combo3.ListIndex = 3 Then
MsgBox "Kota Tidak Boleh Sama", vbInformation, "WARNING!"
End If

If Combo1.Text = "Garuda" Then
Picture1.Picture = LoadPicture(App.Path & "\Garuda2.jpg")
ElseIf Combo1.Text = "Citilink" Then
Picture1.Picture = LoadPicture(App.Path & "\Citilink2.jpg")
ElseIf Combo1.Text = "Lion Air" Then
Picture1.Picture = LoadPicture(App.Path & "\Lion2.jpg")
ElseIf Combo1.Text = "Air Asia" Then
Picture1.Picture = LoadPicture(App.Path & "\Asia2.jpg")
End If
End Sub


Private Sub Option4_Click()
Dim Bayi, Total, Diskon, Anak As Single
Bayi = Val(Text1.Text) * 0.3
Anak = Val(Text1.Text) * 0.5
Select Case Text4.Text
Case Is = 0
 Diskon = 0
Case 1 To 2
 Diskon = 10000
Case 3 To 4
 Diskon = 15000
Case Is >= 5
 Diskon = 20000
End Select
Total = Val(Text2.Text) * Val(Text1.Text) + Bayi * Val(Text3.Text) + Anak * Val(Text4.Text) + (70000 - Diskon)
Text5.Text = Format(Total, "#########")
End Sub

Private Sub Option5_Click()
Dim Bayi, Total, Diskon, Anak As Single
Bayi = Val(Text1.Text) * 0.3
Anak = Val(Text1.Text) * 0.5
Select Case Text4.Text
Case Is = 0
 Diskon = 0
Case 1 To 2
 Diskon = 15000
Case 3 To 4
 Diskon = 20000
Case Is >= 5
 Diskon = 25000
End Select
Total = Val(Text2.Text) * Val(Text1.Text) + Bayi * Val(Text3.Text) + Anak * Val(Text4.Text) + (120000 - Diskon)
Text5.Text = Format(Total, "#########")
End Sub

