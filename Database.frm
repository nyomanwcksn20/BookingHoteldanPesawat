VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Database 
   Caption         =   "Database"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6330
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "<< Back"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next >>"
      Height          =   375
      Left            =   6120
      TabIndex        =   27
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6120
      TabIndex        =   14
      Top             =   3240
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Database.frx":0000
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   5280
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Database.frx":0015
      OLEDBString     =   $"Database.frx":00B1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Tiket_Pesawat"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text12 
      Height          =   405
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text11 
      Height          =   405
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input"
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label12 
      BackColor       =   &H00004080&
      Caption         =   "Harga"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4680
      TabIndex        =   26
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00004080&
      Caption         =   "Bayi"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4680
      TabIndex        =   25
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00004080&
      Caption         =   "Anak-Anak"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4680
      TabIndex        =   24
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00004080&
      Caption         =   "Dewasa"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00004080&
      Caption         =   "Tgl Kembali"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00004080&
      Caption         =   "Tgl Berangkat"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4680
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00004080&
      Caption         =   "Maskapai"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00004080&
      Caption         =   "Ke"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00004080&
      Caption         =   "Dari"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00004080&
      Caption         =   "No. KTP"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004080&
      Caption         =   "No. Tlp"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004080&
      Caption         =   "Nama"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   6375
      Left            =   0
      Picture         =   "Database.frx":014D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset!Nama = Text1.Text
Adodc1.Recordset!Nomer_Telepon = Text2.Text
Adodc1.Recordset!Nomer_KTP = Text3.Text
Adodc1.Recordset!Maskapai = Text4.Text
Adodc1.Recordset!Keberangkatan = Text5.Text
Adodc1.Recordset!Tujuan = Text6.Text
Adodc1.Recordset!Tanggal_Berangkat = Text7.Text
Adodc1.Recordset!Tanggal_Kembali = Text8.Text
Adodc1.Recordset!Dewasa = Text9.Text
Adodc1.Recordset!Anak_Anak = Text10.Text
Adodc1.Recordset!Bayi = Text11.Text
Adodc1.Recordset!Total_Harga = Text12.Text
Adodc1.Recordset.Update
DataGrid1.Refresh
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command3_Click()
Output.Show
Unload Me
End Sub

Private Sub Command4_Click()
If Text8.Text = "-" Then
Tkt_Pswt.Show
Else
Tkt_Pswt2.Show
End If
Unload Me
End Sub
