VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form driver1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15990
   Icon            =   "driver1.frx":0000
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   15990
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   2415
      Left            =   12720
      Picture         =   "driver1.frx":26E8E
      ScaleHeight     =   2355
      ScaleWidth      =   4155
      TabIndex        =   40
      Top             =   240
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   8295
      Left            =   720
      ScaleHeight     =   8235
      ScaleWidth      =   7275
      TabIndex        =   19
      Top             =   1800
      Width           =   7335
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
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   5280
         Width           =   2535
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
         Left            =   3000
         TabIndex        =   15
         ToolTipText     =   "To Exit"
         Top             =   6960
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
         TabIndex        =   5
         Top             =   2280
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
         Left            =   4800
         TabIndex        =   13
         Top             =   6120
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
         Left            =   1320
         TabIndex        =   14
         Top             =   6960
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
         Left            =   1320
         TabIndex        =   2
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
         Left            =   3000
         TabIndex        =   11
         Top             =   6120
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
         Left            =   4800
         TabIndex        =   16
         Top             =   6960
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
         TabIndex        =   12
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
         Height          =   405
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   3
         Top             =   960
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
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text6 
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
         MaxLength       =   12
         TabIndex        =   9
         Top             =   4680
         Width           =   2535
      End
      Begin VB.TextBox Text5 
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
         MaxLength       =   16
         TabIndex        =   8
         Top             =   4080
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
         Top             =   1590
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
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox Text4 
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
         MaxLength       =   10
         TabIndex        =   7
         Top             =   3480
         Width           =   2535
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
         Left            =   840
         TabIndex        =   38
         Top             =   5400
         Width           =   555
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1440
         TabIndex        =   37
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label21 
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
         Left            =   840
         TabIndex        =   36
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver UID"
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
         Top             =   4800
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Licence Id"
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
         TabIndex        =   34
         Top             =   4200
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver ID"
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
         TabIndex        =   33
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Name"
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
         TabIndex        =   32
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Gender"
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
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Address"
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
         Top             =   3000
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Phone no"
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
         Top             =   3600
         Width           =   1395
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   7800
         Width           =   4110
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2040
         TabIndex        =   25
         Top             =   1680
         Width           =   120
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Top             =   2400
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Top             =   3000
         Width           =   120
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Top             =   3600
         Width           =   225
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2400
         TabIndex        =   21
         Top             =   4200
         Width           =   225
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   4800
         Width           =   225
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   15480
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      RecordSource    =   "select * from driver_mst order by b_id,d_id asc"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "driver1.frx":29C80
      Height          =   5895
      Left            =   8520
      TabIndex        =   18
      Top             =   4200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10398
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
      ColumnCount     =   9
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
         DataField       =   "D_ID"
         Caption         =   "DRIVER ID"
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
         DataField       =   "D_NM"
         Caption         =   "DRIVER NAME"
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
         DataField       =   "D_GEN"
         Caption         =   "DRIVER GENDER"
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
         DataField       =   "D_ADDR"
         Caption         =   "DRIVER ADDRESS"
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
         DataField       =   "D_PHNO"
         Caption         =   "DRIVER PHONE NUMBER"
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
         DataField       =   "D_LID"
         Caption         =   "DRIVER LICENCE ID"
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
         DataField       =   "D_UID"
         Caption         =   "DRIVER UID"
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
      BeginProperty Column08 
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
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2039.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2145.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You Can Add, View, Edit And Update The Driver Details"
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
      TabIndex        =   39
      Top             =   1080
      Width           =   6450
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Available Drivers :"
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
      TabIndex        =   17
      Top             =   3600
      Width           =   3765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DRIVER DETAILS"
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
      Width           =   3585
   End
End
Attribute VB_Name = "driver1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
conn
Combo1.AddItem ("Select Bus Id")
Combo1.ListIndex = 0
Combo2.AddItem ("Select Gender")
Combo2.AddItem ("MALE")
Combo2.AddItem ("FEMALE")
Combo2.ListIndex = 0
Combo3.AddItem ("Select Status")
Combo3.AddItem ("TRUE")
Combo3.AddItem ("FALSE")
Combo3.ListIndex = 0
Unload booking1
Unload bus1
Unload conductor1
Unload customer1
Unload report1
Unload route1
Unload stoppage1
Unload bill1
Unload trip1
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
Text2.Enabled = False
End Sub

'for driver id
Private Sub Combo1_click()
conn
sql = "select * from bus_mst where B_id='" + Combo1.Text + "'"
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

'to add new
Private Sub add_Click()
gen
Text1.Locked = True
Text2.Enabled = True
Combo1.Locked = True
Text2.SetFocus
add.Enabled = False
search.Enabled = False
save.Enabled = True
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
Text1.Locked = False
Text1.Text = ""
Text2.Text = ""
Combo2.ListIndex = 0
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo3.ListIndex = 0
add.Enabled = False
save.Enabled = False
search.Enabled = False
update.Enabled = False
Combo1.Locked = False
Text2.Enabled = False
MsgBox "Details Cleared", vbOKOnly, "To Clear"
Else
Exit Sub
End If
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
        u = MsgBox("Please Select A Bus Id.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Driver Id To Find Driver Details.", vbOKOnly, "Driver Id Cannot Be Empty")
        Text1.SetFocus
    ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Driver Name.", vbOKOnly, "Driver Name Cannot Be Empty")
        Text2.SetFocus
    ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Gender.", vbOKOnly, "Gender Cannot Be Empty")
        Combo2.SetFocus
    ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Driver Address.", vbOKOnly, "Driver Address Cannot Be Empty")
        Text3.SetFocus
    ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Driver Phone Number.", vbOKOnly, "Driver Phone Number Cannot Be Empty")
        Text4.SetFocus
    ElseIf Len(Text4.Text) < 10 Then
        u = MsgBox("Not A Valid Phone Number", vbOKOnly, "Enter A valid Phone Number ")
        Text4.SetFocus
    ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Driver Licence Id.", vbOKOnly, "Driver Licence Id Cannot Be Empty")
        Text5.SetFocus
        ElseIf Len(Text5.Text) < 16 Then
        u = MsgBox("Not A Valid Licence Id", vbOKOnly, "Enter A valid Licence Id ")
        Text5.SetFocus
    ElseIf Text6.Text = "" Then
        u = MsgBox("Please Enter Driver UID Id.", vbOKOnly, "Driver Licence UId Cannot Be Empty")
        Text6.SetFocus
        ElseIf Len(Text6.Text) < 12 Then
        u = MsgBox("Not A Valid UID Number ", vbOKOnly, "Enter A valid UID Number ")
        Text6.SetFocus
    ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo3.SetFocus
   Else
      conn
        s = MsgBox("Do You Want To Save Driver Details.", vbQuestion + vbYesNo, "To Save")
            sql = "insert into driver_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + Text2.Text + "','" + Combo2.Text + "','" + Text3.Text + "'," + Text4.Text + ",'" + Text5.Text + "','" + Text6.Text + "','" + Combo3.Text + "')"
    If s = vbYes Then
        c.Execute (sql)
        MsgBox "Record Saved", vbOKOnly, "To Save"
        Else
        Exit Sub
    End If
Adodc1.Refresh
Combo1.Locked = False
Combo1.ListIndex = 0
Combo1.SetFocus
Text1.Locked = False
Text1.Text = " "
Text2.Enabled = False
Text2.Text = " "
Combo2.ListIndex = 0
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Combo3.ListIndex = 0
End If
End Sub


'To  Search
Private Sub search_Click()
conn
Text2.Enabled = True
Text2.Text = ""
Combo2.ListIndex = 0
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo3.ListIndex = 0
    z = MsgBox("Do You Want To Search Driver Details", vbQuestion + vbYesNo, "To Search Driver Details")
    If z = vbYes Then
    Combo1.Locked = True
    Text1.Locked = True
    add.Enabled = False
    search.Enabled = False
    update.Enabled = True
        sql = "select * from driver_mst where d_id='" & Text1.Text & "'"
        Set r = c.Execute(sql)
        If r.EOF = 1 Then
            MsgBox "No matching details found.Please check the Driver Id.", vbOKOnly, "To Search Driver Details"
            Combo1.Locked = False
            Combo1.ListIndex = 0
            Text1.Locked = False
            Text1.Text = ""
            Text2.Enabled = False
            Exit Sub
        End If
MsgBox "Driver Details Found", vbOKOnly, "To Search Driver Details"
Text2.Text = r.Fields(2)
Combo2.Text = r.Fields(3)
Text3.Text = r.Fields(4)
Text4.Text = r.Fields(5)
Text5.Text = r.Fields(6)
Text6.Text = r.Fields(7)
Combo3.Text = r.Fields(8)
End If
End Sub

'To Update
Private Sub update_Click()
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Bus Id.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Drive Id To Find Driver Details.", vbOKOnly, "Driver Id Cannot Be Empty")
        Text1.SetFocus
    ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Driver Name.", vbOKOnly, "Driver Name Cannot Be Empty")
        Text2.SetFocus
    ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Gender.", vbOKOnly, "Gender Cannot Be Empty")
        Combo2.SetFocus
    ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Driver Address.", vbOKOnly, "Driver Address Cannot Be Empty")
        Text3.SetFocus
    ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Driver Phone Number.", vbOKOnly, "Driver Phone Number Cannot Be Empty")
        Text4.SetFocus
    ElseIf Len(Text4.Text) < 10 Then
        u = MsgBox("Not A Valid Phone Number", vbOKOnly, "Enter A valid Phone Number ")
        Text4.SetFocus
    ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Driver Licence Id.", vbOKOnly, "Driver Licence Id Cannot Be Empty")
        Text5.SetFocus
        ElseIf Len(Text5.Text) < 16 Then
        u = MsgBox("Not A Valid Licence Id", vbOKOnly, "Enter A valid Licence Id ")
        Text5.SetFocus
    ElseIf Text6.Text = "" Then
        u = MsgBox("Please Enter Driver UID Id.", vbOKOnly, "Driver Licence UId Cannot Be Empty")
        Text6.SetFocus
        ElseIf Len(Text6.Text) < 12 Then
        u = MsgBox("Not A Valid UID Number ", vbOKOnly, "Enter A valid UID Number ")
        Text6.SetFocus
    ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo3.SetFocus
   Else
      conn
        v = MsgBox("Do You Want To Update Driver Details", vbQuestion + vbYesNo, "To Update")
        sql = "update driver_mst set d_nm='" & Text2.Text & "',d_gen='" & Combo2.Text & "',d_addr='" & Text3.Text & "',d_phno=" & Text4.Text & ",d_lid='" & Text5.Text & "',d_uid=" & Text6.Text & ",status='" & Combo3.Text & "' where(b_id='" & Combo1.Text & "') and (d_id='" & Text1.Text & "')"
    If v = vbYes Then
        Set r = c.Execute(sql)
        MsgBox "Record Updated", vbOKOnly, "To Update"
        Else
        Exit Sub
    End If
Adodc1.Refresh
Combo1.Locked = False
Combo1.ListIndex = 0
Combo1.SetFocus
Text1.Text = " "
Text2.Enabled = False
Text2.Text = " "
Combo2.ListIndex = 0
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Combo3.ListIndex = 0
update.Enabled = False
add.Enabled = False
search.Enabled = False
End If
End Sub


'To Exit
Private Sub exit_Click()
u = MsgBox("Do You Want To Save Changes Before Exiting", vbQuestion + vbYesNoCancel, "To Exit")
If u = vbYes Then
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Bus Id.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Driver Id To Find Driver Details.", vbOKOnly, "Driver Id Cannot Be Empty")
        Text1.SetFocus
    ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Driver Name.", vbOKOnly, "Driver Name Cannot Be Empty")
        Text2.SetFocus
    ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Gender.", vbOKOnly, "Gender Cannot Be Empty")
        Combo2.SetFocus
    ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Driver Address.", vbOKOnly, "Driver Address Cannot Be Empty")
        Text3.SetFocus
    ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Driver Phone Number.", vbOKOnly, "Driver Phone Number Cannot Be Empty")
        Text4.SetFocus
    ElseIf Len(Text4.Text) < 10 Then
        u = MsgBox("Not A Valid Phone Number", vbOKOnly, "Enter A valid Phone Number ")
        Text4.SetFocus
    ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Driver Licence Id.", vbOKOnly, "Driver Licence Id Cannot Be Empty")
        Text5.SetFocus
        ElseIf Len(Text5.Text) < 16 Then
        u = MsgBox("Not A Valid Licence Id", vbOKOnly, "Enter A valid Licence Id ")
        Text5.SetFocus
    ElseIf Text6.Text = "" Then
        u = MsgBox("Please Enter Driver UID Id.", vbOKOnly, "Driver Licence UId Cannot Be Empty")
        Text6.SetFocus
        ElseIf Len(Text6.Text) < 12 Then
        u = MsgBox("Not A Valid UID Number ", vbOKOnly, "Enter A valid UID Number ")
        Text6.SetFocus
    ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo3.SetFocus
   Else
      conn
sql = "insert into driver_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + Text2.Text + "','" + Combo2.Text + "','" + Text3.Text + "'," + Text4.Text + ",'" + Text5.Text + "','" + Text6.Text + "','" + Combo3.Text + "')"
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


'To Generate Report

Private Sub report_Click()
    rptBtn_click = True
    Load report1
    report1.Show
End Sub

'auto genarate driver id
Public Sub gen()
sql = "select max (to_number(SUBSTR(D_id,4,length(D_id)))) from driver_mst"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
Text1.Text = "D0" & "0" & 1
Else
Text1.Text = "D0" & "0" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "D" & "001" & 0) Then
sql = "select max (to_number(SUBSTR(D_id,3,length(D_id)))) from driver_mst"
Set r = c.Execute(sql)
Text1.Text = "D" & "0" & r.Fields(0) + 1
End If
End Sub

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

' For Upper Case
Private Sub Text5_LostFocus()
Text5.Text = UCase(Text5.Text)
Text6.Enabled = True
End Sub

' For Only Number
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 43 Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub
' For Only Number
Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 43 Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

