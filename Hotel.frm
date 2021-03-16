VERSION 5.00
Begin VB.Form Hotel 
   Caption         =   "Booking Kamar Hotel"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10920
   LinkTopic       =   "Form2"
   Picture         =   "Hotel.frx":0000
   ScaleHeight     =   6360
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1560
      TabIndex        =   24
      Top             =   1560
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   360
      Width           =   3735
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Top             =   960
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1560
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   2880
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   4200
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FF80&
      Caption         =   "Reguler"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2160
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080FF80&
      Caption         =   "Deluxe"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   3720
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   4440
      Width           =   3735
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Fasilitas"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Width           =   5055
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080FF80&
         Caption         =   "Tv Kabel dan Internet"
         Height          =   375
         Left            =   2880
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080FF80&
         Caption         =   "Tv Kabel"
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "200000"
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
         Left            =   3120
         TabIndex        =   28
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "150000"
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
         Left            =   840
         TabIndex        =   27
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Kota"
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   6000
      Top             =   720
      Width           =   4575
   End
   Begin VB.Image Picture1 
      Height          =   3375
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Lihat Gambar"
      Height          =   615
      Left            =   6000
      TabIndex        =   22
      Top             =   240
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   5640
      X2              =   5640
      Y1              =   0
      Y2              =   9000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Hotel"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   4320
      Width           =   1005
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Lama Inap"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Kamar"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   4440
      Width           =   1005
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rp."
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   4920
      Width           =   975
   End
End
Attribute VB_Name = "Hotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
Main_Menu.Show
Unload Me
End Sub

Private Sub CmdNext_Click()
Database2.Text5 = Combo2.Text
Database2.Text6 = Combo3.Text
Database2.Text4 = Combo1.Text
Database2.Text7 = List1.Text + List2.Text + List3.Text
Database2.Text8 = Text1.Text
Database2.Text9 = Text2.Text
Database2.Text10 = Text4.Text



Output2.Label2 = Combo2.Text
Output2.Label5 = Combo3.Text
Output2.Label7 = Text2.Text
Output2.Label9 = Text1.Text
Output2.Label12 = List1.Text
Output2.Label13 = List2.Text
Output2.Label14 = List3.Text
Output2.Label15 = Combo1.Text
Output2.Label30 = Text4.Text
If Option1.Enabled = True Then
Output2.Label16 = "Reguler"
ElseIf Option2.Enabled = True Then
Output2.Label16 = "Deluxe"
End If
Database2.Show
Unload Me
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
Combo2.List(0) = "The Ritz-Carlton"
Combo2.List(1) = "Grand Hyatt"
ElseIf Combo1.ListIndex = 1 Then
Combo2.List(0) = "Bumi Surabaya City Resort"
Combo2.List(1) = "Ascott Waterplace"
ElseIf Combo1.ListIndex = 2 Then
Combo2.List(0) = "The Sakala Resort Bali"
Combo2.List(1) = "Kuta Paradiso Hotel"
ElseIf Combo1.ListIndex = 3 Then
Combo2.List(0) = "Grand Nanggroe"
Combo2.List(1) = "Hermes Palace Hotel"
End If
End Sub


Private Sub Combo2_Click()
If Combo2.Text = "The Ritz-Carlton" Then
Picture1.Picture = LoadPicture(App.Path & "\jkt10.jpg")
ElseIf Combo2.Text = "Grand Hyatt" Then
Picture1.Picture = LoadPicture(App.Path & "\jkt20.jpg")
ElseIf Combo2.Text = "Bumi Surabaya City Resort" Then
Picture1.Picture = LoadPicture(App.Path & "\sby10.jpg")
ElseIf Combo2.Text = "Ascott Waterplace" Then
Picture1.Picture = LoadPicture(App.Path & "\sby20.jpg")
ElseIf Combo2.Text = "The Sakala Resort Bali" Then
Picture1.Picture = LoadPicture(App.Path & "\dps10.jpg")
ElseIf Combo2.Text = "Kuta Paradiso Hotel" Then
Picture1.Picture = LoadPicture(App.Path & "\dps20.jpg")
ElseIf Combo2.Text = "Grand Nanggroe" Then
Picture1.Picture = LoadPicture(App.Path & "\ach10.jpg")
ElseIf Combo2.Text = "Hermes Palace Hotel" Then
Picture1.Picture = LoadPicture(App.Path & "\ach20.jpg")
End If
End Sub



Private Sub Form_Load()
Combo1.List(0) = "Jakarta"
Combo1.List(1) = "Surabaya"
Combo1.List(2) = "Bali"
Combo1.List(3) = "Aceh"

Combo3.List(0) = "Single Bed"
Combo3.List(1) = "Double Bed"

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

End Sub


Private Sub Option1_Click()
If Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
Text3.Text = 600000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text3.Text = 800000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text3.Text = 500000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
Text3.Text = 750000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
Text3.Text = 550000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text3.Text = 700000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text3.Text = 600000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
Text3.Text = 750000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
Text3.Text = 600000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text3.Text = 850000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text3.Text = 650000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
Text3.Text = 800000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
Text3.Text = 550000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text3.Text = 650000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text3.Text = 650000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
Text3.Text = 750000
End If

If Combo2.Text = "The Ritz-Carlton" Then
Picture1.Picture = LoadPicture(App.Path & "\jkt11.jpg")
ElseIf Combo2.Text = "Grand Hyatt" Then
Picture1.Picture = LoadPicture(App.Path & "\jkt21.jpg")
ElseIf Combo2.Text = "Bumi Surabaya City Resort" Then
Picture1.Picture = LoadPicture(App.Path & "\sby11.jpg")
ElseIf Combo2.Text = "Ascott Waterplace" Then
Picture1.Picture = LoadPicture(App.Path & "\sby21.jpg")
ElseIf Combo2.Text = "The Sakala Resort Bali" Then
Picture1.Picture = LoadPicture(App.Path & "\dps11.jpg")
ElseIf Combo2.Text = "Kuta Paradiso Hotel" Then
Picture1.Picture = LoadPicture(App.Path & "\dps21.jpg")
ElseIf Combo2.Text = "Grand Nanggroe" Then
Picture1.Picture = LoadPicture(App.Path & "\ach11.jpg")
ElseIf Combo2.Text = "Hermes Palace Hotel" Then
Picture1.Picture = LoadPicture(App.Path & "\ach21.jpg")
End If

End Sub

Private Sub Option2_Click()
If Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
Text3.Text = 800000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text3.Text = 1000000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text3.Text = 700000
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
Text3.Text = 800000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
Text3.Text = 650000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text3.Text = 900000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text3.Text = 700000
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
Text3.Text = 900000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
Text3.Text = 750000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text3.Text = 1050000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text3.Text = 800000
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
Text3.Text = 1050000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
Text3.Text = 900000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
Text3.Text = 900000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
Text3.Text = 800000
ElseIf Combo1.ListIndex = 3 And Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
Text3.Text = 900000
End If

If Combo2.Text = "The Ritz-Carlton" Then
Picture1.Picture = LoadPicture(App.Path & "\jkt12.jpg")
ElseIf Combo2.Text = "Grand Hyatt" Then
Picture1.Picture = LoadPicture(App.Path & "\jkt22.jpg")
ElseIf Combo2.Text = "Bumi Surabaya City Resort" Then
Picture1.Picture = LoadPicture(App.Path & "\sby12.jpg")
ElseIf Combo2.Text = "Ascott Waterplace" Then
Picture1.Picture = LoadPicture(App.Path & "\sby22.jpg")
ElseIf Combo2.Text = "The Sakala Resort Bali" Then
Picture1.Picture = LoadPicture(App.Path & "\dps12.jpg")
ElseIf Combo2.Text = "Kuta Paradiso Hotel" Then
Picture1.Picture = LoadPicture(App.Path & "\dps22.jpg")
ElseIf Combo2.Text = "Grand Nanggroe" Then
Picture1.Picture = LoadPicture(App.Path & "\ach12.jpg")
ElseIf Combo2.Text = "Hermes Palace Hotel" Then
Picture1.Picture = LoadPicture(App.Path & "\ach22.jpg")
End If

End Sub

Private Sub Option3_Click()
inap = Text1.Text
kamar = Text2.Text
harga = Text3.Text
Fasilitas = 150000
subtotal = (inap * harga * kamar)
Total = subtotal + Fasilitas
Text4.Text = Total
End Sub

Private Sub Option4_Click()
inap = Text1.Text
kamar = Text2.Text
harga = Text3.Text
Fasilitas = 200000
subtotal = (inap * harga * kamar)
Total = subtotal + Fasilitas
Text4.Text = Total
End Sub

