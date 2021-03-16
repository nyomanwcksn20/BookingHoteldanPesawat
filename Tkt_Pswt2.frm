VERSION 5.00
Begin VB.Form Tkt_Pswt2 
   Caption         =   "Tiket Pesawat (PP)"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Tkt_Pswt2.frx":0000
   ScaleHeight     =   7470
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5040
      TabIndex        =   34
      Top             =   3840
      Width           =   855
   End
   Begin VB.ListBox List7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5040
      TabIndex        =   33
      Top             =   3120
      Width           =   855
   End
   Begin VB.ListBox List6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      TabIndex        =   32
      Top             =   3840
      Width           =   855
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2640
      TabIndex        =   31
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1680
      TabIndex        =   30
      Top             =   3840
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Bagasi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   15
      Top             =   6480
      Width           =   5655
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080FF80&
         Caption         =   "0 - 5 Kg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H0080FF80&
         Caption         =   "5 Kg >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   16
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "120000"
         Height          =   255
         Left            =   4200
         TabIndex        =   38
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "70000"
         Height          =   255
         Left            =   1560
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   240
      Width           =   4215
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   13
      Text            =   "Combo2"
      Top             =   960
      Width           =   4215
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Text            =   "Combo3"
      Top             =   1680
      Width           =   4215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FF80&
      Caption         =   "Ekonomi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   2400
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080FF80&
      Caption         =   "Bisnis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1680
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2640
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4920
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   5160
      Width           =   4215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5760
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Selanjutnya"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Potongan 70%"
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Potongan 50%"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Pulang / Jam (WIB)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Maskapai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ke"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Tiket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   24
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Berangkat / Jam (WIB)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Dewasa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Harga Rp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   20
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Lihat Gambar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   360
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   7440
   End
   Begin VB.Image Picture1 
      Height          =   3855
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      Height          =   3855
      Left            =   6840
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Anak - Anak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   5160
      Width           =   975
   End
End
Attribute VB_Name = "Tkt_Pswt2"
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
Database.Text8 = List4.Text + List5.Text + List6.Text
Database.Text9 = Text2.Text
Database.Text10 = Text3.Text
Database.Text11 = Text4.Text
Database.Text12 = Text5.Text

Output.Label1 = Combo2.Text
Output.Label2 = Combo3.Text
Output.Label5 = "Kembali     :"
Output.Label7 = Combo1.Text
Output.Label8 = List7.Text
Output.Label9 = List1.Text
Output.Label10 = List2.Text
Output.Label11 = List3.Text
Output.Label12 = List4.Text
Output.Label13 = List5.Text
Output.Label14 = List6.Text
Output.Label17 = Text3.Text
Output.Label20 = Text2.Text
Output.Label21 = Text4.Text
Output.Label22 = List8.Text
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
For b = 2017 To 2050
List6.AddItem b
Next
For h = 1 To 31
List1.AddItem h
Next
For a = 1 To 31
List4.AddItem a
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

List5.List(0) = "Januari"
List5.List(1) = "Februari"
List5.List(2) = "Maret"
List5.List(3) = "April"
List5.List(4) = "Mei"
List5.List(5) = "Juni"
List5.List(6) = "Juli"
List5.List(7) = "Agustus"
List5.List(8) = "September"
List5.List(9) = "Oktober"
List5.List(10) = "November"
List5.List(11) = "Desember"

List7.List(0) = "08:00"
List7.List(1) = "10:00"
List7.List(2) = "12:00"
List7.List(3) = "14:00"
List7.List(4) = "16:00"
List7.List(5) = "18:00"
List7.List(6) = "20:00"
List7.List(7) = "22:00"
List7.List(8) = "24:00"

List8.List(0) = "08:00"
List8.List(1) = "10:00"
List8.List(2) = "12:00"
List8.List(3) = "14:00"
List8.List(4) = "16:00"
List8.List(5) = "18:00"
List8.List(6) = "20:00"
List8.List(7) = "22:00"
List8.List(8) = "24:00"

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

Private Sub List6_Click()
If List4.ListIndex <= List1.ListIndex And List5.ListIndex <= List2.ListIndex And List6.ListIndex <= List3.ListIndex Then
MsgBox "tanggal kembali tidak boleh melebihi tanggal berangkat", vbInformation, "WARNING!"
End If
End Sub

Private Sub Option1_Click()
If Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 1000000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 1200000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 1200000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 1300000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 800000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 850000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1100000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1500000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1700000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 1000000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1450000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1600000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 600000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 650000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 700000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 750000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 500000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 500000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1000000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1300000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1500000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 900000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1250000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1300000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 600000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 650000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 800000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 850000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 450000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 500000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 700000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1000000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1200000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 800000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 900000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1000000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 800000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 850000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 800000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 900000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 550000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 550000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 900000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1000000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1200000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 900000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1100000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1200000 * 2
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
Text1.Text = 1700000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 1650000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 2000000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 2300000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 1300000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 1400000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1500000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1700000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1800000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 1500000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1900000 * 2
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 2000000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 1600000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 1500000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 1800000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 2000000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 900000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 1000000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1300000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1500000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1700000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 1150000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1500000 * 2
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1800000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 800000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 900000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 1100000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 1150000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 700000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 800000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1000000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1100000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1500000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 1000000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1150000 * 2
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1200000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text1.Text = 1150000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text1.Text = 1250000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 2 Then
Text1.Text = 1100000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
Text1.Text = 1200000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 2 Then
Text1.Text = 700000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
Text1.Text = 750000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 3 Then
Text1.Text = 1000000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 3 Then
Text1.Text = 1200000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 2 And Combo3.ListIndex = 3 Then
Text1.Text = 1300000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 0 Then
Text1.Text = 1050000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 1 Then
Text1.Text = 1500000 * 2
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 3 And Combo3.ListIndex = 2 Then
Text1.Text = 1600000 * 2
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

