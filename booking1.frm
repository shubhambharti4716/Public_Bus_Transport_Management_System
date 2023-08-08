VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form booking1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15660
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "booking1.frx":0000
   LinkTopic       =   "Form15"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   15660
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   10080
      Picture         =   "booking1.frx":26E8E
      ScaleHeight     =   2475
      ScaleWidth      =   6795
      TabIndex        =   45
      Top             =   600
      Width           =   6855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   720
      ScaleHeight     =   8715
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   1320
      Width           =   6855
      Begin VB.CommandButton bill 
         Caption         =   "Bill"
         Height          =   375
         Left            =   4920
         TabIndex        =   51
         Top             =   6960
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   6120
         Width           =   2535
      End
      Begin VB.CommandButton clear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2760
         TabIndex        =   46
         ToolTipText     =   "To Exit"
         Top             =   7680
         Width           =   1095
      End
      Begin VB.CommandButton exit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   4560
         TabIndex        =   41
         Top             =   7680
         Width           =   1095
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         Height          =   375
         Left            =   2040
         TabIndex        =   40
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton add 
         Caption         =   "Add New"
         Height          =   375
         Left            =   600
         TabIndex        =   39
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton report 
         Caption         =   "Report"
         Height          =   375
         Left            =   1080
         TabIndex        =   38
         Top             =   7680
         Width           =   1095
      End
      Begin VB.CommandButton update 
         Caption         =   "Update"
         Height          =   375
         Left            =   3480
         TabIndex        =   37
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton search 
         Caption         =   "Search"
         Height          =   375
         Left            =   5400
         TabIndex        =   36
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "booking1.frx":2C094
         Left            =   2640
         List            =   "booking1.frx":2C096
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   1920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   85917697
         CurrentDate     =   43216
         MaxDate         =   44196
         MinDate         =   43216
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   8
         Top             =   4320
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   7
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   6
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   9
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   2520
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   85917697
         CurrentDate     =   43216
         MaxDate         =   44196
         MinDate         =   43216
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   2640
         TabIndex        =   35
         Top             =   4920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   85917697
         CurrentDate     =   43216
         MaxDate         =   44196
         MinDate         =   43216
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs"
         Height          =   210
         Left            =   5280
         TabIndex        =   50
         Top             =   3840
         Width           =   225
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs"
         Height          =   210
         Left            =   5280
         TabIndex        =   49
         Top             =   4440
         Width           =   225
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs"
         Height          =   210
         Left            =   5280
         TabIndex        =   48
         Top             =   3240
         Width           =   225
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   360
         TabIndex        =   43
         Top             =   6240
         Width           =   555
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   960
         TabIndex        =   42
         Top             =   6240
         Width           =   75
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1080
         TabIndex        =   34
         Top             =   240
         Width           =   75
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1440
         TabIndex        =   33
         Top             =   840
         Width           =   75
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2400
         TabIndex        =   32
         Top             =   5640
         Width           =   75
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bus ID"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Id"
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1920
         TabIndex        =   28
         Top             =   5040
         Width           =   75
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1680
         TabIndex        =   27
         Top             =   4440
         Width           =   75
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1920
         TabIndex        =   26
         Top             =   3840
         Width           =   75
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1680
         TabIndex        =   25
         Top             =   3240
         Width           =   75
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2040
         TabIndex        =   24
         Top             =   2640
         Width           =   75
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1920
         TabIndex        =   23
         Top             =   2040
         Width           =   75
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1440
         TabIndex        =   22
         Top             =   1440
         Width           =   75
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
         TabIndex        =   21
         Top             =   8280
         Width           =   4110
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   6
         Left            =   3360
         TabIndex        =   20
         Top             =   8280
         Width           =   90
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   6360
         Y1              =   8160
         Y2              =   8160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Amount"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   3240
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date For Booking"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   2640
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Booking"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Booking ID"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dues Amount"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   4440
         Width           =   1140
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Till Date"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   5040
         Width           =   1485
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total No Of Passenger"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   5640
         Width           =   1950
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   15240
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "select * from booking_mst order by b_id,cu_id,bo_id asc"
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
      Bindings        =   "booking1.frx":2C098
      Height          =   5415
      Left            =   8520
      TabIndex        =   11
      Top             =   4680
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   14737632
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
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
      ColumnCount     =   11
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
         DataField       =   "CU_ID"
         Caption         =   "CUSTOMER ID"
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
      BeginProperty Column03 
         DataField       =   "D_O_B"
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
      BeginProperty Column04 
         DataField       =   "D_F_B"
         Caption         =   "DATE FOR BOOKING"
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
         DataField       =   "TOT_AM"
         Caption         =   "TOTAL  AMOUNT"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "B_T_D"
         Caption         =   "BOOKING TILL DATE"
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
      BeginProperty Column09 
         DataField       =   "T_N_O_P"
         Caption         =   "TOTAL NUMBER OF PASSENGER"
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
      BeginProperty Column10 
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
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1964.976
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   3105.071
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You Can Add, View, Edit And Update The Booking Details"
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
      TabIndex        =   44
      Top             =   840
      Width           =   6675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BOOKING DETAILS"
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
      TabIndex        =   29
      Top             =   120
      Width           =   4005
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Available Bookings :"
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
      TabIndex        =   10
      Top             =   4080
      Width           =   4020
   End
End
Attribute VB_Name = "booking1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim noOfPassenger As Integer
Dim billBtn_click As Boolean

Private Sub bill_Click()
billBtn_click = True
Load bill1
bill1.Show
End Sub


Private Sub DTPicker3_Click()
   DTPicker3.MinDate = DTPicker2.Value
End Sub

Private Sub Form_Load()
DTPicker1.MinDate = Now
DTPicker1.Enabled = False
DTPicker2.MinDate = Now
conn
Combo1.AddItem ("Select Bus Id")
Combo1.ListIndex = 0
Combo2.AddItem ("Select Customer Id")
Combo2.ListIndex = 0
Combo3.AddItem ("Select Status")
Combo3.AddItem ("TRUE")
Combo3.AddItem ("FALSE")
Combo3.ListIndex = 0
Unload bus1
Unload conductor1
Unload customer1
Unload driver1
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
sql = "select cu_id From customer_mst"
Set a = c.Execute(sql)
While a.EOF <> True
Combo2.AddItem a.Fields(0)
a.MoveNext
Wend
add.Enabled = False
save.Enabled = False
update.Enabled = False
search.Enabled = False
DTPicker1.Value = Now
DTPicker2.Value = Now
DTPicker3.Value = Now
Combo2.Enabled = False
Text1.Enabled = False
Text4.Locked = True
End Sub

'for bus id
Private Sub Combo1_click()
conn
   If Combo1.Text = "Select Bus Id" Then
      Exit Sub
   End If
sql = "select * from bus_mst where B_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
noOfPassenger = r.Fields(4)
End Sub

Private Sub Combo1_lostfocus()
If Combo1.ListIndex <> 0 Then
Combo2.Enabled = True
Else
Combo2.ListIndex = 0
Combo2.Enabled = False
End If
End Sub

'TO ADD customer ID
Private Sub Combo2_Click()
sql = "select * from customer_mst where cu_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)
End Sub

Private Sub Combo2_LostFocus()
If Combo2.ListIndex <> 0 Then
Text1.Enabled = True
add.Enabled = True
search.Enabled = True
Else
Combo2.ListIndex = 0
Text1.Enabled = False
add.Enabled = False
search.Enabled = False
End If
End Sub

'To Add New
Private Sub add_Click()
gen
Combo1.Locked = True
Combo2.Locked = True
search.Enabled = False
save.Enabled = True
add.Enabled = False
Text1.Locked = True
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
        u = MsgBox("Please Select Bus Id To Proceed.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Customer Id To Proceed.", vbOKOnly, "Customer Id Cannot Be Empty")
        Combo2.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Booking Id To Find Booking Details.", vbOKOnly, "Booking Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Total Amount To Proceed.", vbOKOnly, "Total Amount Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Advance Amount Paid To Proceed.", vbOKOnly, "Advance Amount")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Re-enter Advance Amount Paid To Find Dues Amount.", vbOKOnly, "Dues Amount")
        Text4.SetFocus
   ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Total Number Of Passenger To Travel.", vbOKOnly, "Total Number Of Passenger")
        Text5.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo3.SetFocus
    Else

    conn
        Dim dt1, dt2, dt3 As String
        dt1 = Format(DTPicker1.Value, "dd/MMM/yyyy")
        dt2 = Format(DTPicker2.Value, "dd/MMM/yyyy")
        dt3 = Format(DTPicker3.Value, "dd/MMM/yyyy")
        s = MsgBox("Do You Want To Save Booking Details.", vbQuestion + vbYesNo, "To Save")
        sql = "insert into booking_mst values ('" + Combo1.Text + "','" + Combo2.Text + "','" + Text1.Text + "','" + dt1 + "','" + dt2 + "','" + Text2.Text + "','" + Text3.Text + "','" + Text4.Text + "','" + dt3 + "'," + Text5.Text + ",'" + Combo3.Text + "')"
    If s = vbYes Then
        c.Execute (sql)
        MsgBox "Record Saved", vbOKOnly, "To Save"
        Else
        Exit Sub
    End If
Adodc1.Refresh
Combo1.Locked = False
Combo2.Locked = False
Combo1.ListIndex = 0
Combo1.SetFocus
Combo2.ListIndex = 0
Text1.Text = " "
Text1.Locked = False
DTPicker1.Value = Now
DTPicker2.Value = Now
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
DTPicker3.Value = Now
Text5.Text = " "
Combo3.ListIndex = 0
save.Enabled = False
add.Enabled = False
search.Enabled = False
update.Enabled = False
End If
End Sub

'To Search
Private Sub search_Click()
conn
DTPicker1.Value = Now
DTPicker2.Value = Now
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
DTPicker3.Value = Now
Text5.Text = " "
Combo3.ListIndex = 0
add.Enabled = False
save.Enabled = False
Combo1.Locked = True
Combo2.Locked = True
z = MsgBox("Do You Want To Search Booking Details", vbQuestion + vbYesNo, "To Search")
sql = "select * from booking_mst where (b_id='" & Combo1.Text & "' and cu_id='" + Combo2.Text & "' and bo_id='" + Text1.Text & "')"
If z = vbYes Then
Set r = c.Execute(sql)
If r.EOF = 1 Then
   MsgBox "No Matching Details Found. Please Check The Bus Id & Customer Id & Booking Id.", vbOKOnly, "To Search"
   Combo1.Locked = False
   Combo2.Locked = False
   Combo1.ListIndex = 0
   Combo2.ListIndex = 0
   Combo1.SetFocus
   Text1.Text = ""
   search.Enabled = False
   Exit Sub
End If
update.Enabled = True
MsgBox "Boking Details Found", vbOKOnly, "To Search"
Combo1.Locked = True
Combo2.Locked = True
Text1.Locked = True
Dim dt As String
dt = r.Fields(3)
'DTPicker1.Value = r.Fields(3)
DTPicker1.Value = dt
Text2.Text = r.Fields(5)
Text3.Text = r.Fields(6)
Text4.Text = r.Fields(7)
DTPicker3.Value = r.Fields(8)
Text5.Text = r.Fields(9)
Combo3.Text = r.Fields(10)
search.Enabled = False
add.Enabled = False
update.Enabled = True
End If
End Sub


Private Sub Text5_LostFocus()
    If Val(Text5.Text) > noOfPassenger + 5 Then
       s = MsgBox("Exceeding Number Of Passengers", vbOKOnly, "Total Number Of Passengers")
       Text5.SetFocus
       Text5.Text = ""
       Exit Sub
    End If
 End Sub

'To Update
Private Sub update_Click()
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select Bus Id To Proceed.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Customer Id To Proceed.", vbOKOnly, "Customer Id Cannot Be Empty")
        Combo2.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Booking Id To Find Booking Details.", vbOKOnly, "Booking Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Total Amount To Proceed.", vbOKOnly, "Total Amount Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Advance Amount Paid To Proceed.", vbOKOnly, "Advance Amount")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Re-enter Advance Amount Paid To Fine Dues Amount.", vbOKOnly, "Dues Amount")
        Text4.SetFocus
   ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Total Number Of Passenger To Travel.", vbOKOnly, "Total Number Of Passenger")
        Text5.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo3.SetFocus
    Else

    conn
        Dim dt1, dt2, dt3 As String
        dt1 = Format(DTPicker1.Value, "dd/MMM/yyyy")
        dt2 = Format(DTPicker2.Value, "dd/MMM/yyyy")
        dt3 = Format(DTPicker3.Value, "dd/MMM/yyyy")
        v = MsgBox("Do You Want To Update Booking Details", vbQuestion + vbYesNo, "To Update")
        sql = "update booking_mst set d_o_b='" & dt1 & "',d_f_b='" & dt2 & "',tot_am='" & Text2.Text & "',adv_am='" & Text3.Text & "',due_am='" & Text3.Text & "',b_t_d='" & dt3 & "',t_n_o_p='" & Text5.Text & "',status='" & Combo3.Text & "' where (b_id='" & Combo1.Text & "')  and (cu_id='" & Combo2.Text & "') and (bo_id='" & Text1.Text & "')"
    If v = vbYes Then
       c.Execute (sql)
        MsgBox "Record Updated", vbOKOnly, "To Update"
        Else
        Exit Sub
    End If
    Adodc1.Refresh
    Combo1.Locked = False
    Combo2.Locked = False
    Combo1.ListIndex = 0
    Combo1.SetFocus
    Combo2.ListIndex = 0
    Combo2.Enabled = False
    Text1.Locked = False
    Text1.Text = ""
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    DTPicker3.Value = Now
    Text5.Text = ""
    Combo3.ListIndex = 0
    update.Enabled = False
    add.Enabled = False
    save.Enabled = False
    search.Enabled = False
  End If
End Sub


Private Sub clear_Click()
If Combo1.Text = "" Then
   s = MsgBox("All Fields Are Already Empty.", vbOKOnly, "Please Fill All The Fields")
   Exit Sub
End If
s = MsgBox("Do you want to Clear Details.", vbQuestion + vbYesNo, "To Clear All Filled Details")
If s = vbYes Then
Combo1.ListIndex = 0
Combo1.SetFocus
Combo2.ListIndex = 0
Combo2.Enabled = False
Text1.Text = ""
Text1.Locked = False
DTPicker1.Value = Now
DTPicker2.Value = Now
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
DTPicker3.Value = Now
Text5.Text = ""
Combo3.ListIndex = 0
add.Enabled = False
save.Enabled = False
search.Enabled = False
update.Enabled = False
MsgBox "Details Cleared", vbOKOnly, "To Clear"
Else
   Exit Sub
End If
End Sub


'To Exit
Private Sub exit_Click()
u = MsgBox("Do You Want To Save Changes Before Exiting", vbQuestion + vbYesNoCancel, "To Exit")
If u = vbYes Then
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select Bus Id To Proceed.", vbOKOnly, "Bus Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select Customer Id To Proceed.", vbOKOnly, "Customer Id Cannot Be Empty")
        Combo2.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Booking Id To Find Booking Details.", vbOKOnly, "Booking Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Total Amount To Proceed.", vbOKOnly, "Total Amount Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Advance Amount Paid To Proceed.", vbOKOnly, "Advance Amount")
        Text3.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Re-enter Advance Amount Paid To Find Dues Amount.", vbOKOnly, "Dues Amount")
        Text4.SetFocus
   ElseIf Text5.Text = "" Then
        u = MsgBox("Please Enter Total Number Of Passenger To Travel.", vbOKOnly, "Total Number Of Passenger")
        Text5.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo3.SetFocus
    Else

    conn
        Dim dt1, dt2, dt3 As String
        dt1 = Format(DTPicker1.Value, "dd/MMM/yyyy")
        dt2 = Format(DTPicker2.Value, "dd/MMM/yyyy")
        dt3 = Format(DTPicker3.Value, "dd/MMM/yyyy")
sql = "insert into booking_mst values ('" + Combo1.Text + "','" + Combo2.Text + "','" + Text1.Text + "','" + dt1 + "','" + dt2 + "','" + Text2.Text + "','" + Text3.Text + "','" + Text4.Text + "','" + dt3 + "'," + Text5.Text + ",'" + Combo3.Text + "')"
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
   If rptBtn_click = True Or mdiBtn_click Or billBtn_click = True Then
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


Public Sub gen()
sql = "select max (to_number(SUBSTR(BO_id,6,length(BO_id)))) from booking_mst"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
Text1.Text = "BO" & "000" & 1
Else
Text1.Text = "BO" & "000" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "BO" & "0001" & 0) Then
sql = "select max (to_number(SUBSTR(BO_id,5,length(BO_id)))) from booking_mst"
Set r = c.Execute(sql)
Text1.Text = "BO" & "00" & r.Fields(0) + 1
End If
End Sub

' For Upper Case
Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

' For Only Number
Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

Private Sub Text3_LostFocus()
If Text2.Text <> "" Then
If Text3.Text <> "" Then
Text4.Text = Text2.Text - Text3.Text
End If
End If
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


