VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form trip1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15705
   Icon            =   "trip1.frx":0000
   LinkTopic       =   "Form13"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   15705
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   11400
      Picture         =   "trip1.frx":26E8E
      ScaleHeight     =   3195
      ScaleWidth      =   5475
      TabIndex        =   27
      Top             =   240
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   8295
      Left            =   720
      ScaleHeight     =   8235
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   1800
      Width           =   6975
      Begin VB.CommandButton clear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   36
         ToolTipText     =   "To Exit"
         Top             =   5520
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3480
         Width           =   2535
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton exit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   24
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   23
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton add 
         Caption         =   "Add New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   22
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton report 
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   21
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton update 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   20
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton search 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   5
         TabIndex        =   1
         ToolTipText     =   "To Add New Trip Id Click Add New Or Enter The Existing Trip Id To Search"
         Top             =   960
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   1560
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   85917697
         CurrentDate     =   43216
         MaxDate         =   44196
         MinDate         =   43216
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   30
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1080
         TabIndex        =   25
         Top             =   3600
         Width           =   75
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bus Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   16
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trip ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trip Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   3000
         Width           =   1440
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   6360
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   6
         Left            =   3360
         TabIndex        =   11
         Top             =   7800
         Width           =   90
      End
      Begin VB.Label label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE- Fields marked with    are mandatory."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   10
         Top             =   7800
         Width           =   4110
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1440
         TabIndex        =   8
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2040
         TabIndex        =   6
         Top             =   3000
         Width           =   90
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   90
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "trip1.frx":2C126
      Height          =   5295
      Left            =   8520
      TabIndex        =   4
      Top             =   4800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9340
      _Version        =   393216
      BackColor       =   14737632
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "B_ID"
         Caption         =   "BUS ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "TR_ID"
         Caption         =   "TRIP ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "TR_DA"
         Caption         =   "TRIP DATE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "SOU_TI"
         Caption         =   "SOURCE TIME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "DES_TI"
         Caption         =   "DESTINATION TIME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "STATUS"
         Caption         =   "STATUS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Locked          =   -1  'True
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   15600
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=MSDAORA.1;Password=abc;User ID=sb;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=abc;User ID=sb;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from trip_mst order by b_id, tr_id asc"
      Caption         =   ""
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You Can Add, View, Edit And Update The Trip Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   18
      Top             =   1080
      Width           =   6210
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Available Trips :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8520
      TabIndex        =   3
      Top             =   4200
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRIP DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      TabIndex        =   17
      Top             =   240
      Width           =   2940
   End
End
Attribute VB_Name = "trip1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
conn
Combo1.AddItem ("Select Bus Id")
Combo1.ListIndex = 0
Combo6.AddItem ("Select Status")
Combo6.AddItem ("TRUE")
Combo6.AddItem ("FALSE")
Combo6.ListIndex = 0
'DTPicker1.MinDate = Now
DTPicker1.Value = Now
Dim i As Integer
Dim h As String
Combo2.AddItem ("HH")
For i = 0 To 23
  If i < 10 Then
     h = "0" & i
     Combo2.AddItem (h)
  Else
     Combo2.AddItem (i)
  End If
Next
Combo2.ListIndex = 0
Combo3.AddItem ("MM")
For i = 0 To 59
  If i < 10 Then
     h = "0" & i
     Combo3.AddItem (h)
  Else
     Combo3.AddItem (i)
  End If
Next
Combo3.ListIndex = 0


Combo4.AddItem ("HH")
For i = 0 To 23
  If i < 10 Then
     h = "0" & i
     Combo4.AddItem (h)
  Else
     Combo4.AddItem (i)
  End If
Next
Combo4.ListIndex = 0
Combo5.AddItem ("MM")
For i = 0 To 59
  If i < 10 Then
     h = "0" & i
     Combo5.AddItem (h)
  Else
     Combo5.AddItem (i)
  End If
Next
Combo5.ListIndex = 0
Unload booking1
Unload bus1
Unload conductor1
Unload customer1
Unload driver1
Unload report1
Unload route1
Unload stoppage1
Unload bill1
HOMEPAGE.Picture3.Visible = False
formopen = 1
sql = "select b_id From bus_mst"
Set r = c.Execute(sql)
While r.EOF <> True
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
add.Enabled = False
save.Enabled = False
update.Enabled = False
search.Enabled = False
Text1.Enabled = False
End Sub

'for bus id
Private Sub Combo1_click()
conn
sql = "select * from bus_mst where B_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
End Sub

Private Sub Combo1_lostfocus()
Text1.Enabled = True
add.Enabled = True
search.Enabled = True
End Sub

'To Add New
Private Sub add_Click()
gen
save.Enabled = True
search.Enabled = False
DTPicker1.SetFocus
End Sub

Private Sub Picture1_Click()
If Combo1.ListIndex <> 0 Then
add.Enabled = True
save.Enabled = False
search.Enabled = True
update.Enabled = False
Else
add.Enabled = False
save.Enabled = False
search.Enabled = False
update.Enabled = False
End If
End Sub

'To Save
Private Sub save_Click()
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Bus Id To Proceed.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Trip Id To Find Stoppage Details.", vbOKOnly, "Trip Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Hours To Proceed.", vbOKOnly, "Select HH")
        Combo2.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select Minutes To Proceed.", vbOKOnly, "Select MM")
        Combo3.SetFocus
   ElseIf Combo4.ListIndex = 0 Then
        u = MsgBox("Please Select Hours To Proceed.", vbOKOnly, "Select HH")
        Combo4.SetFocus
   ElseIf Combo5.ListIndex = 0 Then
        u = MsgBox("Please Select Minutes To Proceed.", vbOKOnly, "Select MM")
        Combo5.SetFocus
   ElseIf Combo6.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo6.SetFocus
    Else

    conn
        Dim userTime1, userTime2, dt As String
        userTime1 = Combo2.Text + ":" + Combo3.Text
        userTime2 = Combo4.Text + ":" + Combo5.Text
        dt = Format(DTPicker1.Value, "dd/MMM/yyyy")
        s = MsgBox("Do You Want To Save Trip Details.", vbQuestion + vbYesNo, "To Save")
            sql = "insert into trip_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + dt + "','" + userTime1 + "','" + userTime2 + "','" + Combo6.Text + "')"
    If s = vbYes Then
        c.Execute (sql)
        MsgBox "Record Saved", vbOKOnly, "To Save"
        Else
        Exit Sub
    End If
Adodc1.Refresh
Combo1.ListIndex = 0
Text1.Text = " "
DTPicker1.Value = Now
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Combo4.ListIndex = 0
Combo5.ListIndex = 0
Combo6.ListIndex = 0
save.Enabled = False
add.Enabled = False
search.Enabled = False
update.Enabled = False
End If
End Sub

'To Search
Private Sub search_Click()
conn
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Combo4.ListIndex = 0
Combo5.ListIndex = 0
Combo6.ListIndex = 0
add.Enabled = False
save.Enabled = False
Combo1.Locked = True
z = MsgBox("Do You Want To Search Trip Details", vbQuestion + vbYesNo, "To Search")
sql = "select * from trip_mst where (b_id='" & Combo1.Text & "' and tr_id='" + Text1.Text & "')"
If z = vbYes Then
Set r = c.Execute(sql)
If r.EOF = True Then
   MsgBox "No Matching Details Found. Please Check The Bus Id & Trip Id.", vbOKOnly, "To Search"
   Combo1.Locked = False
   Combo1.ListIndex = 0
   Combo1.SetFocus
   Text1.Text = ""
   Exit Sub
End If
update.Enabled = True
MsgBox "Stoppage Details Found", vbOKOnly, "To Search"
Text1.Locked = True
DTPicker1.Value = r.Fields(2)
t1 = r.Fields(3)
t1h = Left(t1, 2)
t1m = Right(t1, 2)
Combo2.ListIndex = t1h + 1
Combo3.ListIndex = t1m + 1
t2 = r.Fields(4)
t2h = Left(t2, 2)
t2m = Right(t2, 2)
Combo4.ListIndex = t2h + 1
Combo5.ListIndex = t2m + 1
Combo6.Text = r.Fields(5)
search.Enabled = False
add.Enabled = False
update.Enabled = True
End If
End Sub

'To Update

Private Sub update_Click()
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Bus Id To Proceed.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Trip Id To Find Stoppage Details.", vbOKOnly, "Trip Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Hours To Proceed.", vbOKOnly, "Select HH")
        Combo2.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select Minutes To Proceed.", vbOKOnly, "Select MM")
        Combo3.SetFocus
   ElseIf Combo4.ListIndex = 0 Then
        u = MsgBox("Please Select Hours To Proceed.", vbOKOnly, "Select HH")
        Combo4.SetFocus
   ElseIf Combo5.ListIndex = 0 Then
        u = MsgBox("Please Select Minutes To Proceed.", vbOKOnly, "Select MM")
        Combo5.SetFocus
   ElseIf Combo6.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo6.SetFocus
    Else

    conn
        Dim userTime1, userTime2, dt As String
        userTime1 = Combo2.Text + ":" + Combo3.Text
        userTime2 = Combo4.Text + ":" + Combo5.Text
        dt = Format(DTPicker1.Value, "dd/MMM/yyyy")
        v = MsgBox("Do You Want To Update Trip Details", vbQuestion + vbYesNo, "To Update")
        sql = "update trip_mst set tr_da='" & dt & "',sou_ti='" & userTime1 & "',des_ti='" & userTime2 & "',status='" & Combo6.Text & "' where (b_id='" & Combo1.Text & "') and (tr_id='" & Text1.Text & "')"
    If v = vbYes Then
       c.Execute (sql)
        MsgBox "Record Updated", vbOKOnly, "To Update"
        Else
        Exit Sub
    End If
    Adodc1.Refresh
    Text1.Text = " "
    Combo1.Locked = False
    Combo1.ListIndex = 0
    DTPicker1.Value = Now
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    Combo4.ListIndex = 0
    Combo5.ListIndex = 0
    Combo6.ListIndex = 0
    update.Enabled = False
    End If
End Sub

Private Sub clear_Click()
If Combo1.Text = "" Then
   s = MsgBox("All Fields Are Already Empty.", vbOKOnly, "Please Fill All The Fields")
   Exit Sub
End If
s = MsgBox("Do you want to Clear Details.", vbQuestion + vbYesNo, "To Clear All Filled Details")
If s = vbYes Then
Text1.Text = ""
Combo1.ListIndex = 0
DTPicker1.Value = Now
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Combo4.ListIndex = 0
Combo5.ListIndex = 0
Combo6.ListIndex = 0
End If
Text1.Enabled = False
Combo1.SetFocus
add.Enabled = False
save.Enabled = False
search.Enabled = False
update.Enabled = False
End Sub

Public Sub gen()
sql = "select max (to_number(SUBSTR(Tr_id,4,length(Tr_id)))) from trip_mst where b_id='" & Combo1.Text & "'"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
Text1.Text = "TR" & "0" & 1
Else
Text1.Text = "TR" & "0" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "TR" & "01" & 0) Then
sql = "select max (to_number(SUBSTR(TR_id,3,length(TR_id)))) from trip_mst"
Set r = c.Execute(sql)
Text1.Text = "TR" & r.Fields(0) + 1
End If
End Sub


'To Exit
Private Sub exit_Click()
u = MsgBox("Do You Want To Save Changes Before Exiting", vbQuestion + vbYesNoCancel, "To Exit")
If u = vbYes Then
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Bus Id To Proceed.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Trip Id To Find Stoppage Details.", vbOKOnly, "Trip Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Hours To Proceed.", vbOKOnly, "Select HH")
        Combo2.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select Minutes To Proceed.", vbOKOnly, "Select MM")
        Combo3.SetFocus
   ElseIf Combo4.ListIndex = 0 Then
        u = MsgBox("Please Select Hours To Proceed.", vbOKOnly, "Select HH")
        Combo4.SetFocus
   ElseIf Combo5.ListIndex = 0 Then
        u = MsgBox("Please Select Minutes To Proceed.", vbOKOnly, "Select MM")
        Combo5.SetFocus
   ElseIf Combo6.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo6.SetFocus
    Else

    conn
        Dim userTime1, userTime2, dt As String
        userTime1 = Combo2.Text + ":" + Combo3.Text
        userTime2 = Combo4.Text + ":" + Combo5.Text
        dt = Format(DTPicker1.Value, "dd/MMM/yyyy")
        sql = "insert into trip_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + dt + "','" + userTime1 + "','" + userTime2 + "','" + Combo6.Text + "')"
        c.Execute (sql)
    End If
    Else
    If u = vbNo Then
        Unload Me
    Else
    If u = vbCancel Then
        Exit Sub
    End If
    End If
End If
End Sub

'To Generate Report

Private Sub report_Click()
    rptBtn_click = True
    Load report1
    report1.Show
End Sub

'To Unload
Private Sub Form_QueryUnload(Cancel As Integer, unloadmode As Integer)
   If rptBtn_click = True Or mdiBtn_click = True Then
      HOMEPAGE.Picture3.Visible = True
      Exit Sub
   End If
If MsgBox("Are you sure you want to exit ?", vbQuestion + vbYesNo, "To Exit") = vbNo Then
Cancel = True
Exit Sub
End If
HOMEPAGE.Picture3.Visible = True
End Sub

' For Upper Case
Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

' For only character
Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 64 And KeyAscii < 91 Or KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 20 Or KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 32 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub
' For Upper Case
Private Sub Text2_LostFocus()
Text2.Text = UCase(Text2.Text)
End Sub

