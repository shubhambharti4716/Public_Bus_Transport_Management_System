VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form stoppage1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16125
   Icon            =   "stoppage.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   16125
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   10440
      Picture         =   "stoppage.frx":26E8E
      ScaleHeight     =   3555
      ScaleWidth      =   6435
      TabIndex        =   38
      Top             =   240
      Width           =   6495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   8295
      Left            =   720
      ScaleHeight     =   8235
      ScaleWidth      =   7155
      TabIndex        =   17
      Top             =   1800
      Width           =   7215
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
         Left            =   2760
         TabIndex        =   12
         ToolTipText     =   "To Exit"
         Top             =   6480
         Width           =   1095
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
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4320
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
         Height          =   405
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   3
         ToolTipText     =   "To Add New Stoppage Id Click Add New Or To Search Existing Stoppage Id Write Route Id And Click Search"
         Top             =   1200
         Width           =   2535
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
         Left            =   4560
         TabIndex        =   13
         ToolTipText     =   "To Exit"
         Top             =   6480
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
         Left            =   2760
         TabIndex        =   9
         ToolTipText     =   "To Save Stoppage Details"
         Top             =   5640
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
         Left            =   1080
         TabIndex        =   2
         ToolTipText     =   "To Add New Stoppage Details"
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
         Left            =   1080
         TabIndex        =   11
         ToolTipText     =   "Report Generation"
         Top             =   6480
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
         Left            =   4560
         TabIndex        =   34
         ToolTipText     =   "To Update Existing Stoppage Details"
         Top             =   5640
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
         Left            =   5880
         TabIndex        =   10
         ToolTipText     =   "To Search Existing Stoppage Details"
         Top             =   1200
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
         TabIndex        =   1
         Top             =   600
         Width           =   2535
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
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1830
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
         MaxLength       =   4
         TabIndex        =   5
         Top             =   2460
         Width           =   975
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
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3090
         Width           =   975
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
         MaxLength       =   2
         TabIndex        =   7
         Top             =   3720
         Width           =   975
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
         Left            =   1440
         TabIndex        =   37
         Top             =   4440
         Width           =   90
      End
      Begin VB.Label Label20 
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
         TabIndex        =   35
         Top             =   4440
         Width           =   555
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RS"
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
         TabIndex        =   33
         Top             =   3840
         Width           =   270
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
         TabIndex        =   32
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route Id"
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
         TabIndex        =   14
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stoppage ID"
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
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stoppage Name"
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
         Top             =   1920
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Distance"
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
         Top             =   2520
         Width           =   1260
      End
      Begin VB.Label Label6 
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
         TabIndex        =   28
         Top             =   3240
         Width           =   1320
      End
      Begin VB.Label Label7 
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
         Left            =   840
         TabIndex        =   27
         Top             =   3840
         Width           =   390
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         Left            =   2040
         TabIndex        =   24
         Top             =   1320
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
         Left            =   2280
         TabIndex        =   23
         Top             =   1920
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
         Left            =   2160
         TabIndex        =   22
         Top             =   2520
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
         Left            =   2280
         TabIndex        =   21
         Top             =   3240
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
         Left            =   1320
         TabIndex        =   20
         Top             =   3840
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
         Left            =   1680
         TabIndex        =   19
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label17 
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
         TabIndex        =   18
         Top             =   2520
         Width           =   285
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15360
      Top             =   4320
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
      RecordSource    =   "select * from stoppage_mst order by  r_id,s_id asc"
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
      Bindings        =   "stoppage.frx":3A942
      Height          =   5175
      Left            =   8400
      TabIndex        =   16
      Top             =   4920
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
         DataField       =   "S_ID"
         Caption         =   "STOPPAGE ID"
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
         DataField       =   "S_NM"
         Caption         =   "STOPPAGE NAME"
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
         DataField       =   "DIS"
         Caption         =   "DISTANCE FROM SOURCE (KM)"
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
         DataField       =   "TRA_TI"
         Caption         =   "TRAVELLING TIME"
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
      BeginProperty Column06 
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
         SizeMode        =   1
         BeginProperty Column00 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2924.788
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065.26
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You Can Add, View, Edit And Update The Stoppage Details"
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
      TabIndex        =   36
      Top             =   1080
      Width           =   6795
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Available Stoppages :"
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
      TabIndex        =   15
      Top             =   4320
      Width           =   4155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOPPAGE DETAILS"
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
      TabIndex        =   0
      Top             =   240
      Width           =   4320
   End
End
Attribute VB_Name = "stoppage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_click()
sql = "select * from route_mst where R_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
End Sub

Private Sub Combo1_lostfocus()
If Combo1.ListIndex <> 0 Then
add.Enabled = True
search.Enabled = True
Text1.Enabled = True
Else
add.Enabled = False
search.Enabled = False
Text1.Enabled = True
End If
End Sub


'To Add New
Private Sub add_Click()
gen
Text1.Locked = True
Combo1.Locked = True
Text2.Enabled = True
Text2.SetFocus
add.Enabled = False
save.Enabled = True
search.Enabled = False
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
        u = MsgBox("Please Select A Route Id To Proceed.", vbOKOnly, "Route Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Stoppage Id To Find Stoppage Details.", vbOKOnly, "Stoppage Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Stoppage Name To Proceed.", vbOKOnly, "Stoppage Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Total Distance To Proceed.", vbOKOnly, "Distance Cannot Be Empty")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Travelling Time To Proceed.", vbOKOnly, "Travelling Time Cannot Be Empty")
        Text4.SetFocus
   ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Fare To Proceed.", vbOKOnly, "Fare Cannot Be Empty")
        Text5.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo2.SetFocus
    Else

    conn
        s = MsgBox("Do You Want To Save Stoppage Details.", vbQuestion + vbYesNo, "To Save")
        sql = "insert into stoppage_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + Text2.Text + "'," + Text3.Text + "," + Text4.Text + "," + Text5.Text + ",'" + Combo2.Text + "')"
    If s = vbYes Then
        c.Execute (sql)
        MsgBox "Record Saved", vbOKOnly, "To Save"
    Else
    Exit Sub
End If
Adodc1.Refresh
Combo1.Locked = False
Combo1.ListIndex = 0
Text1.Locked = False
Text1.Text = " "
Text2.Text = " "
Text2.Enabled = False
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Combo2.ListIndex = 0
save.Enabled = False
search.Enabled = False
End If
End Sub

'To Search
Private Sub search_Click()
conn
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Combo2.ListIndex = 0
add.Enabled = False
save.Enabled = False
Combo1.Locked = True
z = MsgBox("Do You Want To Search Stoppage Details", vbQuestion + vbYesNo, "To Search")
sql = "select * from stoppage_mst where (r_id='" & Combo1.Text & "' and s_id='" + Text1.Text & "')"
If z = vbYes Then
Set r = c.Execute(sql)
If r.EOF = 1 Then
   MsgBox "No Matching Details Found. Please Check The Route Id & Stoppage Id.", vbOKOnly, "To Search"
   Combo1.Locked = False
   Combo1.ListIndex = 0
   Text1.Text = ""
   search.Enabled = True
   add.Enabled = True
   Exit Sub
End If
update.Enabled = True
MsgBox "Stoppage Details Found", vbOKOnly, "To Search"
Text1.Locked = True
Text2.Enabled = True
search.Enabled = False
Text2.Text = r.Fields(2)
Text3.Text = r.Fields(3)
Text4.Text = r.Fields(4)
Text5.Text = r.Fields(5)
Combo2.Text = r.Fields(6)
End If
End Sub

'To Update
Private Sub update_Click()
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Route Id To Proceed.", vbOKOnly, "Route Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Stoppage Id To Find Stoppage Details.", vbOKOnly, "Stoppage Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Stoppage Name To Proceed.", vbOKOnly, "Stoppage Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Total Distance To Proceed.", vbOKOnly, "Distance Cannot Be Empty")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Travelling Time To Proceed.", vbOKOnly, "Travelling Time Cannot Be Empty")
        Text4.SetFocus
   ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Fare To Proceed.", vbOKOnly, "Fare Cannot Be Empty")
        Text5.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo2.SetFocus
    Else

    conn
    p = MsgBox("Do You Want To Update Stoppage Details", vbQuestion + vbYesNo, "To Update")
    sql = "update stoppage_mst set s_nm='" & Text2.Text & "',dis=" & Text3.Text & ",tra_ti='" & Text4.Text & "',fare='" & Text5.Text & "',status='" & Combo2.Text & "' where (r_id='" & Combo1.Text & "') and (s_id='" & Text1.Text & "')"
    If p = vbYes Then
    c.Execute (sql)
    MsgBox "Record Updated", vbOKOnly, "To Update"
    Else
    Exit Sub
End If
Adodc1.Refresh
Combo1.Locked = False
Combo1.ListIndex = 0
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Combo2.ListIndex = 0
End If
End Sub

'To Generate Report

Private Sub report_Click()
    rptBtn_click = True
    Load report1
    report1.Show
End Sub

'To Exit
Private Sub exit_Click()
u = MsgBox("Do You Want To Save Changes Before Exiting", vbQuestion + vbYesNoCancel, "To Exit")
If u = vbYes Then
    If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Route Id To Proceed.", vbOKOnly, "Route Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Stoppage Id To Find Stoppage Details.", vbOKOnly, "Stoppage Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Stoppage Name To Proceed.", vbOKOnly, "Stoppage Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Total Distance To Proceed.", vbOKOnly, "Distance Cannot Be Empty")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Travelling Time To Proceed.", vbOKOnly, "Travelling Time Cannot Be Empty")
        Text4.SetFocus
   ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Fare To Proceed.", vbOKOnly, "Fare Cannot Be Empty")
        Text5.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo2.SetFocus
    Else
    conn
        sql = "insert into stoppage_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + Text2.Text + "'," + Text3.Text + "," + Text4.Text + "," + Text5.Text + ",'" + Combo2.Text + "')"
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

'To Clear

Private Sub clear_Click()
If Combo1.ListIndex = 0 Then
   s = MsgBox("All Fields Are Already Empty.", vbOKOnly, "Please Fill All The Fields")
   Exit Sub
End If
s = MsgBox("Do you want to Clear Details.", vbQuestion + vbYesNo, "To Clear All Filled Details")
If s = vbYes Then
Combo1.ListIndex = 0
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo2.ListIndex = 0
add.Enabled = False
save.Enabled = False
search.Enabled = False
update.Enabled = False
Combo1.Locked = False
Text1.Locked = False
Text2.Enabled = False
MsgBox "Details Cleared", vbOKOnly, "To Clear"
Else
Exit Sub
End If
End Sub

Private Sub Form_Load()
conn
Combo1.AddItem ("Select Route Id")
Combo1.ListIndex = 0
Combo2.AddItem ("Select Status")
Combo2.AddItem ("TRUE")
Combo2.AddItem ("FALSE")
Combo2.ListIndex = 0
Unload booking1
Unload bus1
Unload conductor1
Unload customer1
Unload driver1
Unload report1
Unload route1
Unload bill1
Unload trip1
HOMEPAGE.Picture3.Visible = False
formopen = 1
sql = "select R_id From route_mst"
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
Text2.Enabled = False
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

' For Only Number
Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

' For Only Number
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

' For Only Number
Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

Public Sub gen()
sql = "select max (to_number(SUBSTR(S_id,4,length(S_id)))) from stoppage_mst where R_id='" & Combo1.Text & "'"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
Text1.Text = "S" & "00" & 1
Else
Text1.Text = "S" & "00" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "S" & "001" & 0) Then
sql = "select max (to_number(SUBSTR(S_id,3,length(S_id)))) from stoppage_mst"
Set r = c.Execute(sql)
Text1.Text = "S" & "0" & r.Fields(0) + 1
End If
End Sub

