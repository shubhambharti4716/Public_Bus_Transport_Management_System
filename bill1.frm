VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form bill1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16245
   Icon            =   "bill1.frx":0000
   LinkTopic       =   "Form17"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   16245
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   11160
      Picture         =   "bill1.frx":26E8E
      ScaleHeight     =   2715
      ScaleWidth      =   5835
      TabIndex        =   23
      Top             =   240
      Width           =   5895
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
         TabIndex        =   30
         Top             =   6120
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
         Top             =   5280
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
         TabIndex        =   28
         Top             =   5280
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
         TabIndex        =   27
         Top             =   6120
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
         TabIndex        =   26
         Top             =   5280
         Width           =   1095
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
         TabIndex        =   25
         ToolTipText     =   "To Exit"
         Top             =   6120
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
         TabIndex        =   21
         Top             =   960
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
         MaxLength       =   5
         TabIndex        =   6
         Top             =   3960
         Width           =   2535
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
         MaxLength       =   5
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
         MaxLength       =   5
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
         Locked          =   -1  'True
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
         Format          =   83886081
         CurrentDate     =   43216
         MaxDate         =   44196
         MinDate         =   43216
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2760
         TabIndex        =   37
         Top             =   2160
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   83886081
         CurrentDate     =   43216
         MaxDate         =   44196
         MinDate         =   43216
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Billing"
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
         TabIndex        =   39
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label Label6 
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
         TabIndex        =   38
         Top             =   2280
         Width           =   90
      End
      Begin VB.Label Label20 
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
         Left            =   5400
         TabIndex        =   36
         Top             =   4080
         Width           =   225
      End
      Begin VB.Label Label19 
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
         Left            =   5400
         TabIndex        =   35
         Top             =   3480
         Width           =   225
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dues Amount"
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
         TabIndex        =   34
         Top             =   4080
         Width           =   1140
      End
      Begin VB.Label Label22 
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
         TabIndex        =   33
         Top             =   4080
         Width           =   90
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Amount"
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
         TabIndex        =   32
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label21 
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
         TabIndex        =   31
         Top             =   3480
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Top             =   2880
         Width           =   1140
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
         Left            =   5400
         TabIndex        =   20
         Top             =   2880
         Width           =   225
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
         Left            =   1440
         TabIndex        =   18
         Top             =   480
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
         Left            =   1680
         TabIndex        =   17
         Top             =   2880
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
         Left            =   1920
         TabIndex        =   16
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
         Left            =   1080
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         Caption         =   "Booking Id"
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
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Booking"
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
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill ID"
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
         Width           =   525
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
         Top             =   4800
         Width           =   75
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "bill1.frx":29C18
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
         DataField       =   "BO_ID"
         Caption         =   "BOOKING ID"
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
         DataField       =   "BI_ID"
         Caption         =   "BILL ID"
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
         DataField       =   "D_O_BO"
         Caption         =   "DATE OF BOOKING"
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
         DataField       =   "D_O_BI"
         Caption         =   "DATE OF BILLING"
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
         DataField       =   "TOT_AM"
         Caption         =   "TOTAL AMOUNT"
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
         DataField       =   "ADV_AM"
         Caption         =   "ADVANCE AMOUNT"
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
         DataField       =   "DUE_AM"
         Caption         =   "DUES AMOUNT"
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
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1890.142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1574.929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15720
      Top             =   3840
      Visible         =   0   'False
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
      Connect         =   "Provider=MSDAORA.1;Password=abc;User ID=sb;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=abc;User ID=sb;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from bill_mst order by bo_id,bi_id asc"
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
      Caption         =   "You Can Add, View, Edit And Update The Bill Details"
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
      TabIndex        =   22
      Top             =   1080
      Width           =   5490
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Issued Bill:"
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
      Width           =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BILL DETAILS"
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
      TabIndex        =   19
      Top             =   240
      Width           =   2910
   End
End
Attribute VB_Name = "bill1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
conn
Combo1.AddItem ("Select Booking Id")
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
sql = "select bo_id From booking_mst"
Set r = c.Execute(sql)
While r.EOF <> True
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
DTPicker1.Value = Now
DTPicker1.Value = Now
save.Enabled = True
Text1.Enabled = False
Text4.Locked = True
End Sub

'for booking id
Private Sub Combo1_click()
   If Combo1.Text = "Select Booking Id" Then
      Exit Sub
   End If
   sql = "select * from booking_mst where Bo_id='" + Combo1.Text + "'"
   Set r = c.Execute(sql)
   DTPicker1.Value = r.Fields(3)
   DTPicker2.Value = Now
   DTPicker2.Enabled = False
   Text2.Text = r.Fields(5)
   Text3.Text = r.Fields(6)
   Text4.Text = r.Fields(7)
   gen
End Sub

Private Sub Combo1_lostfocus()
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
'gen
'Combo1.Locked = True
'Text1.Locked = True
'DTPicker1.SetFocus
'save.Enabled = True
'add.Enabled = False
'search.Enabled = False
'update.Enabled = False
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
        u = MsgBox("Please Select A Booking Id To Proceed.", vbOKOnly, " Booking Id")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Bill Id To Find Bill Details.", vbOKOnly, "Bill Id")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Total Amount To Proceed.", vbOKOnly, "Total Amount")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Advance Amount To Proceed.", vbOKOnly, "Advance Amount")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Re-enter Advance Amount To Calculate Dues Amount To Proceed.", vbOKOnly, "Dues Amount")
        Text4.SetFocus
    Else
    conn
        Dim dt1, dt2 As String
        dt1 = Format(DTPicker1.Value, "dd/MMM/yyyy")
        dt2 = Format(DTPicker2.Value, "dd/MMM/yyyy")
        s = MsgBox("Do You Want To Save Bill Details.", vbQuestion + vbYesNo, "To Save")
        sql = "insert into bill_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + dt1 + "', '" + dt2 + "','" + Text2.Text + "','" + Text3.Text + "','" + Text4.Text + "')"
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
DTPicker2.Value = Now
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
Text1.Locked = False
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
add.Enabled = False
save.Enabled = False
z = MsgBox("Do You Want To Search Bill Details", vbQuestion + vbYesNo, "To Search")
sql = "select * from bill_mst where (bo_id ='" & Combo1.Text & "' and bi_id='" & Text1.Text & "')"
If z = vbYes Then
Set r = c.Execute(sql)
If r.EOF = 1 Then
   MsgBox "No Matching Details Found. Please Check The Booking Id And Bill Id.", vbOKOnly, "To Search"
   Combo1.Locked = False
   Combo1.ListIndex = 0
   Text1.Text = ""
   search.Enabled = True
   Exit Sub
   End If
   update.Enabled = True
   MsgBox "Ticket Details Found", vbOKOnly, "To Search"
   DTPicker1.Value = r.Fields(2)
   DTPicker2.Value = r.Fields(3)
   Text2.Text = r.Fields(4)
   Text3.Text = r.Fields(5)
   Text4.Text = r.Fields(6)
End If
End Sub

Private Sub Text3_LostFocus()
If Text2.Text <> "" Then
If Text3.Text <> "" Then
Text4.Text = Text2.Text - Text3.Text
End If
End If
End Sub

'To Update

Private Sub update_Click()
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Booking Id To Proceed.", vbOKOnly, " Booking Id")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Bill Id To Find Bill Details.", vbOKOnly, "Bill Id")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Total Amount To Proceed.", vbOKOnly, "Total Amount")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Advance Amount To Proceed.", vbOKOnly, "Advance Amount")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Re-enter Advance Amount To Calculate Dues Amount To Proceed.", vbOKOnly, "Dues Amount")
        Text4.SetFocus
    Else
        conn
        Dim dt1, dt2 As String
        
        dt = Format(DTPicker1.Value, "dd/MMM/yyyy")
        p = MsgBox("Do You Want To Update Bill Details", vbQuestion + vbYesNo, "To Update")
            sql = "update bill_mst set d_o_bo='" & dt1 & "',d_o_bi='" & dt2 & "',tot_am= '" & Text2.Text & "',adv_am='" & Text3.Text & "',due_am=" & Text4.Text & " where (bo_id='" & Combo1.Text & "') and (bi_id='" & Text1.Text & "')"
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
update.Enabled = False
End If
End Sub

'To Clear

Private Sub clear_Click()
conn
Combo1.ListIndex = 0
Combo1.Locked = False
Text1.Text = ""
Text1.Locked = False
DTPicker1.Value = Now
DTPicker2.Value = Now
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
    rptBtn_click = True
    Load report1
    report1.Show
End Sub


Public Sub gen()
sql = "select max (to_number(SUBSTR(BI_id,4,length(BI_id)))) from bill_mst where bo_id='" & Combo1.Text & "'"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
Text1.Text = "BI" & "00000" & 1
Else
Text1.Text = "BI" & "00000" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "BI" & "00001" & 0) Then
sql = "select max (to_number(SUBSTR(BI_id,3,length(BI_id)))) from trip_mst"
Set r = c.Execute(sql)
Text1.Text = "BI" & "0000" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "BI" & "00010" & 0) Then
sql = "select max (to_number(SUBSTR(BI_id,3,length(BI_id)))) from trip_mst"
Set r = c.Execute(sql)
Text1.Text = "BI" & "000" & r.Fields(0) + 1
End If
End Sub


'To Exit
Private Sub exit_Click()
u = MsgBox("Do You Want To Save Changes Before Exiting", vbQuestion + vbYesNoCancel, "To Exit")
If u = vbYes Then
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Booking Id To Proceed.", vbOKOnly, " Booking Id")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Bill Id To Find Bill Details.", vbOKOnly, "Bill Id")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Total Amount To Proceed.", vbOKOnly, "Total Amount")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Advance Amount To Proceed.", vbOKOnly, "Advance Amount")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Re-enter Advance Amount To Calculate Dues Amount To Proceed.", vbOKOnly, "Dues Amount")
        Text4.SetFocus
    Else
    conn
        Dim dt1, dt2 As String
        dt1 = Format(DTPicker1.Value, "dd/MMM/yyyy")
        dt2 = Format(DTPicker2.Value, "dd/MMM/yyyy")
sql = "insert into bill_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + dt1 + "', '" + dt2 + "','" + Text2.Text + "','" + Text3.Text + "','" + Text4.Text + "')"
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

Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

' For Only Number
Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 43 Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub
' For Only Number
Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 43 Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

' For Only Number
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 43 Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

