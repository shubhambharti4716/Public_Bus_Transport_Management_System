VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form route1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15690
   Icon            =   "route.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   15690
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   3495
      Left            =   10200
      Picture         =   "route.frx":26E8E
      ScaleHeight     =   3435
      ScaleWidth      =   6675
      TabIndex        =   40
      Top             =   120
      Width           =   6735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15240
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "select * from route_mst order by r_id asc"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "route.frx":2ECD1
      Height          =   5175
      Left            =   8400
      TabIndex        =   34
      Top             =   4800
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      BackColor       =   14737632
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "R_ID"
         Caption         =   "ROUTE ID"
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
         DataField       =   "R_NM"
         Caption         =   "ROUTE NAME"
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
         DataField       =   "SOU_PO"
         Caption         =   "SOURCE POINT"
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
         DataField       =   "DES_PO"
         Caption         =   "DESTINATION POINT"
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
         DataField       =   "DIS"
         Caption         =   "TOTAL DISTANCE"
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
         DataField       =   "TRA_TI"
         Caption         =   "TRAVELLING TIME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "AVG_SP"
         Caption         =   "AVERAGE SPEED"
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
      BeginProperty Column07 
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
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2280.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
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
         Left            =   2520
         TabIndex        =   13
         ToolTipText     =   "To Exit"
         Top             =   6480
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
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Select Status Of Route From dropdown"
         Top             =   4680
         Width           =   2535
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
         Left            =   4320
         TabIndex        =   11
         ToolTipText     =   "To Update Existing Route Details"
         Top             =   5640
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
         Left            =   840
         TabIndex        =   12
         ToolTipText     =   "Report Generation"
         Top             =   6480
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
         Left            =   5760
         TabIndex        =   14
         ToolTipText     =   "Search By Route Id"
         Top             =   360
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
         Height          =   405
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   2
         ToolTipText     =   "To Add New Route Id Click Add New Or To Search Existing Route Id Write Route Id And Click Search"
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         DataField       =   "AVG_SP"
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
         Left            =   3120
         MaxLength       =   6
         TabIndex        =   8
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox Text6 
         DataField       =   "TRA_TI"
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
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   7
         ToolTipText     =   "Total Travelling Time In Minutes"
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox Text2 
         DataField       =   "R_NM"
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
         Left            =   3120
         MaxLength       =   49
         TabIndex        =   3
         ToolTipText     =   "Route Name"
         Top             =   990
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         DataField       =   "SOU_PO"
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
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   4
         ToolTipText     =   "Route Source"
         Top             =   1620
         Width           =   2535
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
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   5
         ToolTipText     =   "Route Destination"
         Top             =   2250
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         DataField       =   "DIS"
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
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   6
         ToolTipText     =   "Total Distance In KM"
         Top             =   2880
         Width           =   975
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
         Left            =   840
         TabIndex        =   1
         ToolTipText     =   "To Add New Route"
         Top             =   5640
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
         Left            =   2520
         TabIndex        =   10
         ToolTipText     =   "To Save Route Details"
         Top             =   5640
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
         Left            =   4320
         TabIndex        =   15
         ToolTipText     =   "To Exit"
         Top             =   6480
         Width           =   1095
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
         Height          =   195
         Left            =   2160
         TabIndex        =   39
         Top             =   3600
         Width           =   90
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
         Left            =   2280
         TabIndex        =   38
         Top             =   3000
         Width           =   90
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
         Left            =   1440
         TabIndex        =   37
         Top             =   4800
         Width           =   90
      End
      Begin VB.Label Label18 
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
         Left            =   840
         TabIndex        =   36
         Top             =   4800
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Average Speed"
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
         Left            =   840
         TabIndex        =   31
         Top             =   4200
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Travelling Time"
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
         Left            =   840
         TabIndex        =   30
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route ID"
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
         Left            =   840
         TabIndex        =   29
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route Name"
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
         Left            =   840
         TabIndex        =   28
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route Source"
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
         Left            =   840
         TabIndex        =   27
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route Destination"
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
         Left            =   840
         TabIndex        =   26
         Top             =   2400
         Width           =   1545
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route Distance"
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
         Left            =   840
         TabIndex        =   25
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   6360
         Y1              =   7680
         Y2              =   7680
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
         Index           =   6
         Left            =   3360
         TabIndex        =   24
         Top             =   7800
         Width           =   90
      End
      Begin VB.Label label10 
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
         TabIndex        =   23
         Top             =   7800
         Width           =   4110
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
         Left            =   1680
         TabIndex        =   22
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
         Left            =   1920
         TabIndex        =   21
         Top             =   1080
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
         TabIndex        =   20
         Top             =   1800
         Width           =   90
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM"
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
         Left            =   4200
         TabIndex        =   19
         Top             =   3000
         Width           =   285
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MIN"
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
         Left            =   4200
         TabIndex        =   18
         Top             =   3600
         Width           =   360
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
         Left            =   2400
         TabIndex        =   17
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM/HR"
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
         Left            =   4200
         TabIndex        =   16
         Top             =   4200
         Width           =   645
      End
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Available Routes :"
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
      Left            =   8400
      TabIndex        =   32
      Top             =   4200
      Width           =   3720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You Can Add, View, Edit And Update The Route Details"
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
      TabIndex        =   35
      Top             =   1080
      Width           =   6435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ROUTE DETAILS"
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
      TabIndex        =   33
      Top             =   240
      Width           =   3465
   End
End
Attribute VB_Name = "route1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ch As Integer

'Add New

Private Sub add_Click()
gen
Text1.Locked = True
Text2.Enabled = True
Text2.SetFocus
add.Enabled = False
save.Enabled = True
update.Enabled = False
search.Enabled = False
update.Enabled = False
End Sub

'To Save

Private Sub save_Click()
   If Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Route Id To Find Route Details.", vbOKOnly, "Route Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Route Name.", vbOKOnly, "Route Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Route Source Name.", vbOKOnly, "Route Source Name Cannot Be Empty")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Route Destination Name.", vbOKOnly, "Route Destination Name Cannot Be Empty")
        Text4.SetFocus
   ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Route Total Distance.", vbOKOnly, "Route Distance Cannot Be Empty")
        Text5.SetFocus
   ElseIf Text6.Text = "" Then
        u = MsgBox("Please Enter Route Total Travelling Time.", vbOKOnly, "Total Travelling Time Cannot Be Empty")
        Text6.SetFocus
   ElseIf Text7.Text = "" Then
        u = MsgBox("Please Re-enter the Total Travelling Time.", vbOKOnly, "Average Speed")
        Text6.Text = ""
        Text6.SetFocus
   ElseIf Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo1.SetFocus
   Else
      conn
      v = MsgBox("Do You Want To Save Route Details", vbQuestion + vbYesNo, "To Save")
      sql = "insert into route_mst values ('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "','" + Text4.Text + "'," + Text5.Text + ",'" + Text6.Text + "'," + Text7.Text + ",'" + Combo1.Text + "')"
    If v = vbYes Then
      c.Execute (sql)
      MsgBox "Record Saved", vbOKOnly, "To Save"
    Else
      Exit Sub
    End If
      Adodc1.Refresh
      Text1.Text = ""
      Text2.Text = ""
      Text3.Text = ""
      Text4.Text = ""
      Text5.Text = ""
      Text6.Text = ""
      Text7.Text = ""
      Combo1.ListIndex = 0
      Text1.Locked = False
      save.Enabled = False
      add.Enabled = True
      search.Enabled = True
  End If
End Sub

' Auto generate

Public Sub gen()
    sql = "select max (to_number(SUBSTR(R_id,4,length(R_id)))) from route_mst"
    Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
    Text1.Text = "R" & "00" & 1
Else
    Text1.Text = "R" & "00" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "R" & "001" & 0) Then
    sql = "select max (to_number(SUBSTR(R_id,3,length(R_id)))) from route_mst"
    Set r = c.Execute(sql)
    Text1.Text = "R" & "0" & r.Fields(0) + 1
End If
End Sub

'TO UPDATE

Private Sub update_Click()
  If Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Route Id To Find Route Details.", vbOKOnly, "Route Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Route Name.", vbOKOnly, "Route Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Route Source Name.", vbOKOnly, "Route Source Name Cannot Be Empty")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Route Destination Name.", vbOKOnly, "Route Destination Name Cannot Be Empty")
        Text4.SetFocus
   ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Route Total Distance.", vbOKOnly, "Route Distance Cannot Be Empty")
        Text5.SetFocus
   ElseIf Text6.Text = "" Then
        u = MsgBox("Please Enter Route Total Travelling Time.", vbOKOnly, "Total Travelling Time Cannot Be Empty")
        Text6.SetFocus
   ElseIf Text7.Text = "" Then
        u = MsgBox("Please Re-enter the Total Travelling Time.", vbOKOnly, "Average Speed")
        Text6.Text = ""
        Text6.SetFocus
   ElseIf Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo1.SetFocus
   Else
      conn
        v = MsgBox("Do You Want To Update Route Details", vbQuestion + vbYesNo, "To Update")
        sql = "update route_mst set r_nm='" & Text2.Text & "',sou_po='" & Text3.Text & "',des_po='" & Text4.Text & "',dis=" & Text5.Text & ",tra_ti='" & Text6.Text & "',avg_sp=" & Text7.Text & ",status='" & Combo1.Text & "' where r_id='" & Text1.Text & "'"
    If v = vbYes Then
       Set r = c.Execute(sql)
        MsgBox "Record Updated", vbOKOnly, "To Update"
        Else
        Exit Sub
    End If
         Adodc1.Refresh
         Text1.Text = " "
         Text2.Text = " "
         Text3.Text = " "
         Text4.Text = " "
         Text5.Text = " "
         Text6.Text = " "
         Text7.Text = " "
         Combo1.ListIndex = 0
         add.Enabled = True
         save.Enabled = False
         update.Enabled = False
         Text1.Locked = False
  End If
End Sub

'To Generate Report

Private Sub report_Click()
    rptBtn_click = True
    Load report1
    report1.Show
End Sub

'To exit

Private Sub exit_Click()
   If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text6.Text = "" And Text7.Text = "" And Combo1.ListIndex = 0 Then
      Unload Me
      Exit Sub
   Else
      GoTo validate
      Exit Sub
   End If
validate:
u = MsgBox("Do You Want To Save Changes Before Exiting", vbQuestion + vbYesNoCancel, "To Exit")
If u = vbYes Then
    If Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Route Id To Find Route Details.", vbOKOnly, "Route Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Route Name.", vbOKOnly, "Route Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Route Source Name.", vbOKOnly, "Route Source Name Cannot Be Empty")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Route Destination Name.", vbOKOnly, "Route Destination Name Cannot Be Empty")
        Text4.SetFocus
   ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Route Total Distance.", vbOKOnly, "Route Distance Cannot Be Empty")
        Text5.SetFocus
   ElseIf Text6.Text = "" Then
        u = MsgBox("Please Enter Route Total Travelling Time.", vbOKOnly, "Total Travelling Time Cannot Be Empty")
        Text6.SetFocus
   ElseIf Text7.Text = "" Then
        u = MsgBox("Please Re-enter the Total Travelling Time.", vbOKOnly, "Average Speed")
        Text6.Text = ""
        Text6.SetFocus
   ElseIf Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo1.SetFocus
   Else
    conn
      sql = "insert into route_mst values ('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "','" + Text4.Text + "'," + Text5.Text + ",'" + Text6.Text + "'," + Text7.Text + ",'" + Combo1.Text + "')"
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

'To Search

Private Sub search_Click()
conn
    Text2.Text = " "
    Text3.Text = " "
    Text4.Text = " "
    Text5.Text = " "
    Text6.Text = " "
    Text7.Text = " "
    Combo1.ListIndex = 0
    z = MsgBox("Do You Want To Search Route Details", vbQuestion + vbYesNo, "To Search")
If z = vbYes Then
sql = "select * from route_mst where r_id='" & Text1.Text & "'"
Set r = c.Execute(sql)
If r.EOF = True Then
   MsgBox "No Matching Details Found. Please Check The Route Id.", vbOKOnly, "To Search"
   Text1.Text = ""
   Text1.Locked = False
   Text1.SetFocus
   update.Enabled = False
   Exit Sub
End If
Text1.Locked = True
Text2.Enabled = True
add.Enabled = False
save.Enabled = False
update.Enabled = True
MsgBox "Route Details Found", vbOKOnly, "To Search"
Text2.Text = r.Fields(1)
Text3.Text = r.Fields(2)
Text4.Text = r.Fields(3)
Text5.Text = r.Fields(4)
Text6.Text = r.Fields(5)
Text7.Text = r.Fields(6)
Combo1.Text = r.Fields(7)
End If
End Sub

' To Clear

Private Sub clear_Click()
If Text1.Text = "" Then
   s = MsgBox("All Fields Are Already Empty.", vbOKOnly, "Please Fill All The Fields")
   Exit Sub
End If
s = MsgBox("Do you want to Clear Details.", vbQuestion + vbYesNo, "To Clear All Filled Details")
If s = vbYes Then
Text1.Text = ""
Text1.Locked = False
Text1.SetFocus
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.ListIndex = 0
add.Enabled = True
save.Enabled = False
search.Enabled = True
update.Enabled = False
MsgBox "Details Cleared", vbOKOnly, "To Clear"
Else
   Exit Sub
End If
End Sub

'At Load Time

Private Sub Form_Load()
conn
Combo1.AddItem ("Select Status")
Combo1.AddItem ("TRUE")
Combo1.AddItem ("FALSE")
Combo1.ListIndex = 0
Unload booking1
Unload bus1
Unload conductor1
Unload customer1
Unload driver1
Unload report1
Unload stoppage1
Unload bill1
Unload trip1
HOMEPAGE.Picture2.Visible = True
HOMEPAGE.Picture3.Visible = False
check_for_activeform
formopen = 1
Text2.Enabled = False
Text7.Locked = True
save.Enabled = False
update.Enabled = False
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
If Text1.Text = "" And ch <> 1 Then
u = MsgBox("Please Click Add New Or Enter Route Id To Find Route Details.", vbOKOnly, "Route Id Cannot Be Empty")
add.SetFocus
End If
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
If (KeyAscii > 64 And KeyAscii < 91 Or KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 20 Or KeyAscii = 8 Or KeyAscii = 32 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub
' For Upper Case
Private Sub Text3_LostFocus()
Text3.Text = UCase(Text3.Text)
End Sub
' For only character
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii > 64 And KeyAscii < 91 Or KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 20 Or KeyAscii = 8 Or KeyAscii = 32 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub
' For Upper Case
Private Sub Text4_LostFocus()
Text4.Text = UCase(Text4.Text)
End Sub
' For Only Number
Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub
' For Only Number
Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

Private Sub Text6_LostFocus()
If Text5.Text = "" Then
Text7.Text = ""
Else
If Text6.Text = "" Then
Text7.Text = ""
Else
If Text5.Text <> "" Then
If Text6.Text <> "" Then
Text7.Text = Text5.Text / (Text6.Text / 60)
End If
End If
End If
End If
End Sub

' For Only Number
Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

