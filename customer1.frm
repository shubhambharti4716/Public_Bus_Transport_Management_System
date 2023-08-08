VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form customer1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15855
   Icon            =   "customer1.frx":0000
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   15855
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   3975
      Left            =   10200
      Picture         =   "customer1.frx":26E8E
      ScaleHeight     =   3915
      ScaleWidth      =   6675
      TabIndex        =   29
      Top             =   240
      Width           =   6735
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
      Begin VB.CommandButton booking 
         Caption         =   "Booking"
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
         TabIndex        =   33
         ToolTipText     =   "To Exit"
         Top             =   5760
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   32
         ToolTipText     =   "Select Status from dropdown"
         Top             =   3480
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
         Left            =   2520
         TabIndex        =   31
         ToolTipText     =   "To Exit"
         Top             =   5760
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   30
         ToolTipText     =   "Select Status from dropdown"
         Top             =   2280
         Width           =   2535
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
         Left            =   5640
         TabIndex        =   27
         Top             =   360
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
         Left            =   2520
         TabIndex        =   26
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
         Left            =   2520
         TabIndex        =   25
         Top             =   4920
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
         Left            =   840
         TabIndex        =   24
         Top             =   4920
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
         TabIndex        =   23
         Top             =   5760
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
         Left            =   4320
         TabIndex        =   22
         Top             =   4920
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
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   1
         Top             =   360
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
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   2
         Top             =   960
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
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1680
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
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2880
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
         Left            =   600
         TabIndex        =   21
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
         Left            =   1200
         TabIndex        =   20
         Top             =   3600
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         Left            =   600
         TabIndex        =   18
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Left            =   600
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Age"
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
         Left            =   600
         TabIndex        =   16
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Gender"
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
         Left            =   600
         TabIndex        =   15
         Top             =   2400
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Phone no"
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
         Left            =   600
         TabIndex        =   14
         Top             =   3000
         Width           =   1665
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         Height          =   210
         Left            =   1800
         TabIndex        =   11
         Top             =   480
         Width           =   75
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
         Height          =   210
         Left            =   2040
         TabIndex        =   10
         Top             =   1080
         Width           =   75
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " *"
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
         Left            =   1800
         TabIndex        =   9
         Top             =   1800
         Width           =   135
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
         Height          =   210
         Left            =   2160
         TabIndex        =   8
         Top             =   2400
         Width           =   75
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " *"
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
         Left            =   2280
         TabIndex        =   7
         Top             =   3000
         Width           =   135
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   15480
      Top             =   5040
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
      RecordSource    =   "select * from customer_mst order by  cu_id asc"
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
      Bindings        =   "customer1.frx":2F7E5
      Height          =   4455
      Left            =   8400
      TabIndex        =   6
      Top             =   5640
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7858
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
      BeginProperty Column01 
         DataField       =   "CU_NM"
         Caption         =   "CUSTOMER NAME"
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
         DataField       =   "CU_AGE"
         Caption         =   "CUSTOMER AGE"
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
         DataField       =   "CU_GEN"
         Caption         =   "CUSTOMER GENDER"
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
         DataField       =   "CU_PHNO"
         Caption         =   "CUSTOMER PHONE NUMBER"
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
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2025.071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2700.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You Can Add, View, Edit And Update The Customer Details"
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
      TabIndex        =   28
      Top             =   1080
      Width           =   6870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER REGISTRATION"
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
      Width           =   5865
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Available Customers :"
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
      TabIndex        =   5
      Top             =   5040
      Width           =   4260
   End
End
Attribute VB_Name = "customer1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub booking_Click()
Load booking1
booking1.Show
End Sub

Private Sub Form_Load()
conn
Combo1.AddItem ("Select Gender")
Combo1.AddItem ("MALE")
Combo1.AddItem ("FEMALE")
Combo1.ListIndex = 0
Combo2.AddItem ("Select Status")
Combo2.AddItem ("TRUE")
Combo2.AddItem ("FALSE")
Combo2.ListIndex = 0
Unload booking1
Unload bus1
Unload conductor1
Unload driver1
Unload report1
Unload route1
Unload stoppage1
Unload bill1
Unload trip1
HOMEPAGE.Picture3.Visible = False
formopen = 1
save.Enabled = False
update.Enabled = False
Text2.Enabled = False
End Sub

'To Add New
Private Sub add_Click()
gen
Text2.Enabled = True
Text2.SetFocus
add.Enabled = False
search.Enabled = False
save.Enabled = True
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
If Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Customer Id To Find Route Details.", vbOKOnly, "Customer Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Customer Name.", vbOKOnly, "Customer Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Customer Age.", vbOKOnly, "Customer Age Cannot Be Empty")
        Text3.SetFocus
    ElseIf Val(Text3.Text) < 18 Then
      Dim r As Variant
      r = MsgBox("Only Persons Whose Age Is 18 Or Above Can Be Our Customer.", vbOKOnly, "PBTMS says.....")
      Text3.Text = ""
      Text3.SetFocus
   ElseIf Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select Gender.", vbOKOnly, "Gender Cannot Be Empty")
        Combo1.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Customer Phone Number.", vbOKOnly, "Phone Number Cannot Be Empty")
        Text4.SetFocus
        ElseIf Len(Text4.Text) < 10 Then
        u = MsgBox("Not A Valid Phone Number", vbOKOnly, "Enter A valid Phone Number ")
        Text4.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select A VAlid Status.", vbOKOnly, "For Status")
        Combo2.SetFocus
    Else
        conn
        v = MsgBox("Do You Want To Save Customer Details", vbQuestion + vbYesNo, "To Save")
        sql = "insert into customer_mst values ('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "','" + Combo1.Text + "'," + Text4.Text + ",'" + Combo2.Text + "')"
    If v = vbYes Then
        c.Execute (sql)
MsgBox "Record Saved", vbOKOnly, "To Save"
    Else
      Exit Sub
    End If
Adodc1.Refresh
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Combo1.ListIndex = 0
Text4.Text = " "
Combo2.ListIndex = 0
save.Enabled = False
add.Enabled = True
search.Enabled = True
Text2.Enabled = False
End If
End Sub

'To Search
Private Sub search_Click()
conn
Text2.Enabled = True
Text2.Text = ""
Text3.Text = ""
Combo1.ListIndex = 0
Text4.Text = ""
Combo2.ListIndex = 0
z = MsgBox("Do you want to search Customer Details", vbQuestion + vbYesNo, "To Search Customer Details")
If z = vbYes Then
sql = "select * from customer_mst where cu_id='" & Text1.Text & "'"
Set r = c.Execute(sql)
If r.EOF = 1 Then
   MsgBox "No Matching Details Found. Please Check The Customer Id.", vbOKOnly, "To Search"
   Combo1.ListIndex = 0
   Combo2.ListIndex = 0
   Text1.Text = ""
   Text2.Enabled = False
   search.Enabled = True
    add.Enabled = True
    update.Enabled = False
   Exit Sub
End If
MsgBox "Customer Details Found", vbOKOnly, "To Search"
Text2.Enabled = True
search.Enabled = False
add.Enabled = False
update.Enabled = True
Text2.Text = r.Fields(1)
Text3.Text = r.Fields(2)
Combo1.Text = r.Fields(3)
Text4.Text = r.Fields(4)
Combo2.Text = r.Fields(5)
End If
End Sub


'To Update
Private Sub update_Click()
If Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Customer Id To Find Route Details.", vbOKOnly, "Customer Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Customer Name.", vbOKOnly, "Customer Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Customer Age.", vbOKOnly, "Customer Age Cannot Be Empty")
        Text3.SetFocus
   ElseIf Val(Text3.Text) < 18 Then
      Dim r As Variant
      r = MsgBox("Only Persons Whose Age Is 18 Or Above Can Be Our Customer.", vbOKOnly, "PBTMS says.....")
      Text3.Text = ""
      Text3.SetFocus
   ElseIf Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select Gender.", vbOKOnly, "Gender Cannot Be Empty")
        Combo1.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Customer Phone Number.", vbOKOnly, "Phone Number Cannot Be Empty")
        Text4.SetFocus
        ElseIf Len(Text4.Text) < 10 Then
        u = MsgBox("Not A Valid Phone Number", vbOKOnly, "Enter A valid Phone Number ")
        Text4.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "For Status")
        Combo2.SetFocus
    Else
        conn
            v = MsgBox("Do You Want To Update Customer Details", vbQuestion + vbYesNo, "To Update")
            sql = "update customer_mst set cu_nm='" & Text2.Text & "',cu_age=" & Text3.Text & ",cu_gen='" & Combo1.Text & "',cu_phno=" & Text4.Text & ",status='" & Combo2.Text & "' where cu_id='" & Text1.Text & "'"
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
Combo1.ListIndex = 0
Text4.Text = " "
Combo2.ListIndex = 0
Text2.Enabled = False
update.Enabled = False
add.Enabled = True
search.Enabled = True
End If
End Sub

Private Sub clear_Click()
If Text1.Text = "" Then
   s = MsgBox("All Fields Are Already Empty.", vbOKOnly, "Please Fill All The Fields")
   Exit Sub
End If
s = MsgBox("Do you want to Clear Details.", vbQuestion + vbYesNo, "To Clear All Filled Details")
If s = vbYes Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.ListIndex = 0
Text4.Text = ""
Combo2.ListIndex = 0
add.Enabled = True
save.Enabled = False
search.Enabled = True
update.Enabled = False
Text2.Enabled = False
MsgBox "Details Cleared", vbOKOnly, "To Clear"
Else
   Exit Sub
End If
End Sub


'To Generate Report

Private Sub report_Click()
    rptBtn_click = True
    Load report1
    report1.Show
End Sub


Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub



'To Exit
Private Sub exit_Click()
u = MsgBox("Do You Want To Save Changes Before Exiting", vbQuestion + vbYesNoCancel, "To Exit")
If u = vbYes Then
If Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Customer Id To Find Route Details.", vbOKOnly, "Customer Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Customer Name.", vbOKOnly, "Customer Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Customer Age.", vbOKOnly, "Customer Age Cannot Be Empty")
        Text3.SetFocus
    ElseIf Val(Text3.Text) < 18 Then
      Dim r As Variant
      r = MsgBox("Only Persons Whose Age Is 18 Or Above Can Be Our Customer.", vbOKOnly, "PBTMS says.....")
      Text3.Text = ""
      Text3.SetFocus
   ElseIf Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select Gender.", vbOKOnly, "Gender Cannot Be Empty")
        Combo1.SetFocus
   ElseIf Text4.Text = "" Then
        u = MsgBox("Please Enter Customer Phone Number.", vbOKOnly, "Phone Number Cannot Be Empty")
        Text4.SetFocus
        ElseIf Len(Text4.Text) < 10 Then
        u = MsgBox("Not A Valid Phone Number", vbOKOnly, "Enter A valid Phone Number ")
        Text4.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Select A VAlid Status.", vbOKOnly, "For Status")
        Combo2.SetFocus
    Else
        conn
sql = "insert into customer_mst values ('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "','" + Combo1.Text + "'," + Text4.Text + ",'" + Combo2.Text + "')"
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

Public Sub gen()
sql = "select max (to_number(SUBSTR(CU_id,6,length(CU_id)))) from customer_mst"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
Text1.Text = "CU" & "000" & 1
Else
Text1.Text = "CU" & "000" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "CU" & "001" & 0) Then
sql = "select max (to_number(SUBSTR(CU_id,5,length(CU_id)))) from customer_mst"
Set r = c.Execute(sql)
Text1.Text = "CU" & "00" & r.Fields(0) + 1
End If
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

' For only character
Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii > 64 And KeyAscii < 91 Or KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 20 Or KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 32 Then
keyasci = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub
' For Upper Case
Private Sub Text5_LostFocus()
Text5.Text = UCase(Text5.Text)
End Sub
