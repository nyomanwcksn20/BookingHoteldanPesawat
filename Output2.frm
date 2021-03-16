VERSION 5.00
Begin VB.Form Output2 
   Caption         =   "Booking"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   LinkTopic       =   "Form2"
   ScaleHeight     =   6015
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Selesai"
      Height          =   375
      Left            =   5040
      TabIndex        =   25
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boking Kamar Hotel"
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   6015
      Begin VB.Label Label16 
         Height          =   375
         Left            =   5040
         TabIndex        =   16
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label14 
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Check In     :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Hari"
         Height          =   255
         Left            =   5280
         TabIndex        =   10
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Lama Inap   :"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   4920
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Jmlh Kamar :"
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Type           :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Nama Hotel :"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Rp."
      Height          =   375
      Left            =   1200
      TabIndex        =   26
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label27 
      Height          =   255
      Left            =   1320
      TabIndex        =   24
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label28 
      Height          =   375
      Left            =   1320
      TabIndex        =   23
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label31 
      Caption         =   "Harga   :"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label30 
      Height          =   255
      Left            =   1680
      TabIndex        =   21
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label29 
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label26 
      Caption         =   "No KTP:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label25 
      Caption         =   "No Tlp  :"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "Nama   :"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "Output2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Main_Menu.Show
Unload Me
End Sub
