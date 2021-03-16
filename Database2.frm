VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Database2 
   Caption         =   "Form2"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15330
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   15330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "<< Back"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next >>"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6720
      TabIndex        =   13
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input"
      Height          =   495
      Left            =   8040
      TabIndex        =   12
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   4320
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
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
      Connect         =   $"Database2.frx":0000
      OLEDBString     =   $"Database2.frx":009C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Booking_Hotel"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Database2.frx":0138
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1720
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
   Begin VB.Label Label7 
      BackColor       =   &H00004080&
      Caption         =   "Harga"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5280
      TabIndex        =   23
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00004080&
      Caption         =   "Jumlah Kamar"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00004080&
      Caption         =   "Lama Inap"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5280
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00004080&
      Caption         =   "Tanggal"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00004080&
      Caption         =   "Tipe Kamar"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5280
      TabIndex        =   19
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00004080&
      Caption         =   "Nama Hotel"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00004080&
      Caption         =   "Kota"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004080&
      Caption         =   "No. KTP"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00004080&
      Caption         =   "No. Tlp"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004080&
      Caption         =   "Nama"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   5175
      Left            =   0
      Picture         =   "Database2.frx":014D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "Database2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset!Nama = Text1.Text
Adodc1.Recordset!Nomer_Telepon = Text2.Text
Adodc1.Recordset!Nomer_KTP = Text3.Text
Adodc1.Recordset!Kota = Text4.Text
Adodc1.Recordset!Nama_Hotel = Text5.Text
Adodc1.Recordset!Tipe_Kamar = Text6.Text
Adodc1.Recordset!Tanggal = Text7.Text
Adodc1.Recordset!Lama_Inap = Text8.Text
Adodc1.Recordset!Jumlah_Kamar = Text9.Text
Adodc1.Recordset!Total_Harga = Text10.Text
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
Hotel.Show
Unload Me
End Sub
