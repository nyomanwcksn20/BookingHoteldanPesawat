VERSION 5.00
Begin VB.Form Output 
   Caption         =   "Tiket"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Selesai"
      Height          =   375
      Left            =   7560
      TabIndex        =   32
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiket Pesawat"
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   8775
      Begin VB.Label Label23 
         Height          =   255
         Left            =   7920
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label22 
         Height          =   255
         Left            =   5280
         TabIndex        =   22
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label17 
         Height          =   255
         Left            =   6720
         TabIndex        =   21
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Anak      :"
         Height          =   255
         Left            =   5880
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Dari   :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Ke     : "
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Berangkat :"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label11 
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label14 
         Height          =   255
         Left            =   4560
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "Dewasa  :"
         Height          =   255
         Left            =   5880
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Bayi        :"
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label21 
         Height          =   255
         Left            =   6720
         TabIndex        =   3
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   1
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label20 
         Height          =   375
         Left            =   6720
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Label Label32 
      Caption         =   "Rp."
      Height          =   375
      Left            =   1440
      TabIndex        =   33
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label31 
      Caption         =   "Harga   :"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label30 
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label29 
      Height          =   255
      Left            =   1440
      TabIndex        =   29
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label28 
      Height          =   375
      Left            =   1440
      TabIndex        =   28
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label27 
      Height          =   255
      Left            =   1440
      TabIndex        =   27
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label26 
      Caption         =   "No KTP:"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label25 
      Caption         =   "No Tlp  :"
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "Nama   :"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Output"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Main_Menu.Show
Unload Me
End Sub
