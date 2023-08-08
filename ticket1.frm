VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ticket1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15630
   Icon            =   "ticket1.frx":0000
   LinkTopic       =   "Form17"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   15630
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   10320
      Picture         =   "ticket1.frx":26E8E
      ScaleHeight     =   2475
      ScaleWidth      =   6675
      TabIndex        =   35
      Top             =   240
      Width           =   6735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   8295
      Left            =   720
      ScaleHeight     =   8235
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   1800
      Width           =   7095
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
         TabIndex        =   38
         Top             =   2160
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
         TabIndex        =   37
         Top             =   2160
         Width           =   855
      End
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
         Top             =   6000
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
         TabIndex        =   33
         Top             =   960
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
         TabIndex        =   32
         Top             =   5160
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
         TabIndex        =   31
         Top             =   6000
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
         TabIndex        =   30
         Top             =   5160
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
         TabIndex        =   29
         Top             =   5160
         Width           =   1095
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
         TabIndex        =   28
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         DataField       =   "DES_PO"
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
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         DataField       =   "DES_PO"
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
         MaxLength       =   19
         TabIndex        =   5
         Top             =   3360
         Width           =   2535
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
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         DataField       =   "DES_PO"
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
         MaxLength       =   19
         TabIndex        =   4
         Top             =   2760
         Width           =   2535
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
         MaxLength       =   9
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   1560
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90243073
         CurrentDate     =   43216
         MaxDate         =   44196
         MinDate         =   43216
      End
      Begin VB.Label Label22 
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
         TabIndex        =   39
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3840
         TabIndex        =   27
         Top             =   4080
         Width           =   225
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fare"
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
         Top             =   4080
         Width           =   390
      End
      Begin VB.Label Label19 
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
         Left            =   960
         TabIndex        =   25
         Top             =   4080
         Width           =   90
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Stoppage"
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
         TabIndex        =   24
         Top             =   3480
         Width           =   1110
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
         Left            =   1680
         TabIndex        =   23
         Top             =   3480
         Width           =   90
      End
      Begin VB.Label Label15 
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
         Left            =   1080
         TabIndex        =   21
         Top             =   480
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
         Left            =   1920
         TabIndex        =   20
         Top             =   2880
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
         Left            =   1920
         TabIndex        =   19
         Top             =   2280
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
         Left            =   2040
         TabIndex        =   18
         Top             =   1680
         Width           =   90
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
         Left            =   1440
         TabIndex        =   17
         Top             =   1080
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
         TabIndex        =   16
         Top             =   7800
         Width           =   4110
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
         TabIndex        =   15
         Top             =   7800
         Width           =   90
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   6360
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Label Label7 
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
         TabIndex        =   14
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Stoppage"
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
         Top             =   2880
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Departure Time"
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
         Top             =   2280
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Journey"
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
         TabIndex        =   11
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket ID"
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
         TabIndex        =   10
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   9
         Top             =   3600
         Width           =   75
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ticket1.frx":2E06F
      Height          =   5535
      Left            =   8760
      TabIndex        =   8
      Top             =   4560
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9763
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
      ColumnCount     =   7
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
         DataField       =   "TI_ID"
         Caption         =   "TICKET ID"
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
         DataField       =   "D_O_J"
         Caption         =   "DATE OF JOURNEY"
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
         DataField       =   "DEP_TIME"
         Caption         =   "DEPARTURE TIME"
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
         DataField       =   "FROM_STO"
         Caption         =   "FROM STOPPAGE"
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
         DataField       =   "TO_STO"
         Caption         =   "TO STOPPAGE"
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
      BeginProperty Column06 
         DataField       =   "FARE"
         Caption         =   "FARE"
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
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1874.835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1874.835
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2234.835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   824.882
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15720
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Connect         =   "Provider=MSDAORA.1;Password=abc;User ID=alok;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=abc;User ID=alok;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from ticket_mst order by b_id,ti_id asc"
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
      Caption         =   "You Can Add, View, Edit And Update The Ticket Details"
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
      TabIndex        =   34
      Top             =   1080
      Width           =   6495
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List OF  Issued Ticket :"
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
      Left            =   8760
      TabIndex        =   7
      Top             =   3960
      Width           =   3450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TICKET Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   720
      TabIndex        =   22
      Top             =   240
      Width           =   3525
   End
End
Attribute VB_Name = "ticket1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
conn
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
Combo1.AddItem ("Select Bus Id")
Combo1.ListIndex = 0
Unload booking1
Unload bus1
Unload conductor1
Unload customer1
Unload driver1
Unload report1
Unload route1
Unload stoppage1
Unload trip1
HOMEPAGE.Picture3.Visible = False
formopen = 1
sql = "select b_id From bus_mst"
Set r = c.Execute(sql)
While r.EOF <> True
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
DTPicker1.Value = Now
add.Enabled = False
save.Enabled = False
search.Enabled = False
update.Enabled = False
Text1.Enabled = False
End Sub

'for bus id
Private Sub Combo1_Click()
conn
sql = "select * from bus_mst where B_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
End Sub

Private Sub Combo1_LostFocus()
If Combo1.ListIndex <> 0 Then
Text1.Enabled = True
add.Enabled = True
search.Enabled = True
Else
Text1.Enabled = False
add.Enabled = False
search.Enabled = False
End If
End Sub

'To Add New
Private Sub add_Click()
gen
Combo1.Locked = True
Text1.Locked = True
DTPicker1.SetFocus
save.Enabled = True
add.Enabled = False
search.Enabled = False
update.Enabled = False
End Sub

'To Save
Private Sub save_Click()
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Bus Id To Proceed.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Ticket Id To Find Stoppage Details.", vbOKOnly, "Ticket Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Hours To Proceed.", vbOKOnly, "Select HH")
        Combo2.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select Minutes To Proceed.", vbOKOnly, "Select MM")
        Combo3.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter From Stoppage Name To Proceed.", vbOKOnly, "From Stoppage")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter To Stoppage To Proceed.", vbOKOnly, "To Stoppage")
        Text3.SetFocus
   ElseIf Text4.Text = 0 Then
        u = MsgBox("Please Enter Fare.", vbOKOnly, "Fare")
        Text4.SetFocus
    Else
    conn
        Dim userTime1, dt As String
        userTime1 = Combo2.Text + ":" + Combo3.Text
        dt = Format(DTPicker1.Value, "dd/MMM/yyyy")
        s = MsgBox("Do You Want To Save Ticket Details.", vbQuestion + vbYesNo, "To Save")
        sql = "insert into ticket_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + dt + "', '" + userTime1 + "','" + Text2.Text + "','" + Text3.Text + "'," + Text4.Text + ")"
    If s = vbYes Then
        c.Execute (sql)
        MsgBox "Record Saved", vbOKOnly, "To Save"
        Else
        Exit Sub
    End If
Adodc1.Refresh
Combo1.Locked = False
Combo1.ListIndex = 0
Text1.Text = " "
Text1.Enabled = False
Text1.Locked = False
DTPicker1.Value = Now
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
save.Enabled = False
add.Enabled = False
End If
End Sub

'To Search'
Private Sub search_Click()
conn
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Text1.Locked = False
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
add.Enabled = False
save.Enabled = False
z = MsgBox("Do You Want To Search Ticket Details", vbQuestion + vbYesNo, "To Search")
sql = "select * from ticket_mst where (b_id ='" & Combo1.Text & "' and ti_id='" & Text1.Text & "')"
If z = vbYes Then
Set r = c.Execute(sql)
If r.EOF = True Then
   MsgBox "No Matching Details Found. Please Check The Bus Id And Ticket Id.", vbOKOnly, "To Search"
   Combo1.Locked = False
   Combo1.ListIndex = 0
   Text1.Text = ""
   search.Enabled = True
   Exit Sub
   End If
   update.Enabled = True
   MsgBox "Ticket Details Found", vbOKOnly, "To Search"
   DTPicker1.Value = r.Fields(2)
   t1 = r.Fields(3)
   t1h = Left(t1, 2)
   t1m = Right(t1, 2)
   Combo2.ListIndex = t1h + 1
   Combo3.ListIndex = t1m + 1
   Text2.Text = r.Fields(4)
   Text3.Text = r.Fields(5)
   Text4.Text = r.Fields(6)
End If
End Sub

Private Sub update_Click()
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Bus Id To Proceed.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Ticket Id To Find Stoppage Details.", vbOKOnly, "Ticket Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Hours To Proceed.", vbOKOnly, "Select HH")
        Combo2.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select Minutes To Proceed.", vbOKOnly, "Select MM")
        Combo3.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter From Stoppage Name To Proceed.", vbOKOnly, "From Stoppage")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter To Stoppage To Proceed.", vbOKOnly, "To Stoppage")
        Text3.SetFocus
   ElseIf Text4.Text = 0 Then
        u = MsgBox("Please Enter Fare.", vbOKOnly, "Fare")
        Text4.SetFocus
    Else
        conn
        Dim userTime1, dt As String
        userTime1 = Combo2.Text + ":" + Combo3.Text
        dt = Format(DTPicker1.Value, "dd/MMM/yyyy")
        p = MsgBox("Do You Want To Update Ticket Details", vbQuestion + vbYesNo, "To Update")
            sql = "update ticket_mst set d_o_j='" & dt & "',dep_time= '" & Combo2.Text & ":" & Combo3.Text & "',from_sto='" & Text2.Text & "',to_sto='" & Text3.Text & "',fare=" & Text4.Text & " where (b_id='" & Combo1.Text & "') and (ti_id='" & Text1.Text & "')"
    If p = vbYes Then
        Set a = c.Execute(sql)
        MsgBox "Record Updated", vbOKOnly, "To Update"
        Else
        Exit Sub
End If
Adodc1.Refresh
Combo1.Locked = False
Combo1.ListIndex = 0
Combo1.SetFocus
Text1.Text = " "
Text1.Enabled = False
DTPicker1.Value = Now
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Combo2.ListIndex = 0
Combo3.ListIndex = 0
update.Enabled = False
End If
End Sub



Private Sub clear_Click()
conn
Combo1.ListIndex = 0
Combo1.Locked = False
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Text1.Text = ""
Text1.Enabled = True
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
add.Enabled = True
save.Enabled = False
search.Enabled = True
update.Enabled = False
Combo1.SetFocus
End Sub


'To Generate Report

Private Sub report_Click()
    Report_std.Show
    Report_std.Frame1.Visible = False
    Report_std.Frame2.Visible = False
    Report_std.Frame3.Visible = False
    Report_std.Frame4.Visible = False
    Report_std.Frame5.Visible = False
    Report_std.Frame6.Visible = False
    Report_std.Frame7.Visible = False
    Report_std.Frame8.Visible = False
    Report_std.Frame9.Visible = True
    Report_std.Frame9.Enabled = True
End Sub


Public Sub gen()
sql = "select max (to_number(SUBSTR(TI_id,4,length(TI_id)))) from ticket_mst where b_id='" & Combo1.Text & "'"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
Text1.Text = "TI" & "00000" & 1
Else
Text1.Text = "TI" & "00000" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "TI" & "00001" & 0) Then
sql = "select max (to_number(SUBSTR(TI_id,3,length(TI_id)))) from trip_mst"
Set r = c.Execute(sql)
Text1.Text = "TI" & "0000" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "TI" & "00010" & 0) Then
sql = "select max (to_number(SUBSTR(TI_id,3,length(TI_id)))) from trip_mst"
Set r = c.Execute(sql)
Text1.Text = "TI" & "000" & r.Fields(0) + 1
End If
End Sub


'To Exit
Private Sub exit_Click()
Y = MsgBox("Do You Want To Exit", vbQuestion + vbYesNo, "To Exit")
If Y = vbYes Then
Unload Me
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Y = MsgBox(" You Are Exiting This Window", vbOKOnly, "Exit")
HOMEPAGE.Picture3.Visible = True
End Sub




Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

' For only character
Private Sub Text2_KeyPress(KeyAscii As Integer)
add.Enabled = False
If (KeyAscii > 64 And KeyAscii < 91 Or KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 20 Or KeyAscii = 8 Or KeyAscii = 32 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub
' For Upper Case
Private Sub Text2_LostFocus()
Text2.Text = UCase(Text2.Text)
End Sub
' For only character
Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii > 64 And KeyAscii < 91 Or KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 20 Or KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 32 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub
' For Upper Case
Private Sub Text3_LostFocus()
Text3.Text = UCase(Text3.Text)
End Sub
' For Only Number
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 43 Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

