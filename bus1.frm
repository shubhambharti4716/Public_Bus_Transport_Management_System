VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form bus1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16275
   Icon            =   "bus1.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   16275
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   10200
      Picture         =   "bus1.frx":26E8E
      ScaleHeight     =   2235
      ScaleWidth      =   6675
      TabIndex        =   31
      Top             =   240
      Width           =   6735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "bus1.frx":2CE2A
      Height          =   5895
      Left            =   8520
      TabIndex        =   14
      Top             =   4080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10398
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
      ColumnCount     =   6
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
      BeginProperty Column02 
         DataField       =   "B_NM"
         Caption         =   "BUS NAME"
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
         DataField       =   "T_O_B"
         Caption         =   "TYPE OF BUS"
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
         DataField       =   "T_N_O_S"
         Caption         =   "TOTAL NUMBER OF SEAT"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2445.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   8295
      Left            =   720
      ScaleHeight     =   8235
      ScaleWidth      =   7155
      TabIndex        =   15
      Top             =   1800
      Width           =   7215
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
         TabIndex        =   6
         Top             =   3960
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
         Left            =   2760
         TabIndex        =   11
         ToolTipText     =   "To Exit"
         Top             =   6120
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
         TabIndex        =   4
         Top             =   2760
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
         Left            =   4560
         TabIndex        =   9
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
         Left            =   1080
         TabIndex        =   10
         Top             =   6120
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
         TabIndex        =   28
         Top             =   5280
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
         TabIndex        =   7
         Top             =   5280
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
         Left            =   4560
         TabIndex        =   12
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
         Left            =   5880
         TabIndex        =   8
         Top             =   1320
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
         Top             =   1320
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
         MaxLength       =   20
         TabIndex        =   3
         Top             =   2040
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
         MaxLength       =   2
         TabIndex        =   5
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label8 
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
         Left            =   1440
         TabIndex        =   32
         Top             =   4080
         Width           =   75
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
         TabIndex        =   29
         Top             =   4080
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bus ID"
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
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bus Name"
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
         Top             =   2160
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type Of Bus"
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
         Top             =   2880
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total No Of Seat"
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
         TabIndex        =   24
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         TabIndex        =   23
         Top             =   720
         Width           =   750
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         Left            =   1560
         TabIndex        =   20
         Top             =   1440
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
         Left            =   1920
         TabIndex        =   19
         Top             =   2160
         Width           =   75
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
         Height          =   210
         Left            =   2040
         TabIndex        =   18
         Top             =   2880
         Width           =   75
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
         Left            =   2400
         TabIndex        =   17
         Top             =   3480
         Width           =   75
      End
      Begin VB.Label Label17 
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
         TabIndex        =   16
         Top             =   720
         Width           =   75
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   15360
      Top             =   3360
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
      RecordSource    =   "select * from bus_mst order by r_id,b_id asc"
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You Can Add, View, Edit And Update The Bus Details"
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
      TabIndex        =   30
      Top             =   1080
      Width           =   6165
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Available Buses :"
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
      TabIndex        =   13
      Top             =   3360
      Width           =   3540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUS DETAILS"
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
      Width           =   2850
   End
End
Attribute VB_Name = "bus1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_click()
sql = "select * from route_mst where R_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
End Sub

Private Sub Combo1_lostfocus()
add.Enabled = True
search.Enabled = True
Text1.Enabled = True
End Sub

'To Add New
Private Sub add_Click()
gen
Text2.Enabled = True
Text2.SetFocus
save.Enabled = True
search.Enabled = False
update.Enabled = False
End Sub

Private Sub add_LostFocus()
add.Enabled = False
Combo1.Locked = True
Text1.Locked = True
Text2.Enabled = True
Text2.SetFocus
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
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Route Id To Proceed.", vbOKOnly, "Route Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Bus Id To Find Bus Details.", vbOKOnly, "Bus Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Bus Name To Proceed.", vbOKOnly, "Bus Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Enter Type Of Bus To Proceed.", vbOKOnly, "Bus Type Cannot Be Empty")
        Combo2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Total Number Of Seat To Proceed.", vbOKOnly, "Total Number Of Seats Cannot Be Empty")
        Text3.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo3.SetFocus
    Else

    conn
    s = MsgBox("Do You Want To Save Bus Details.", vbQuestion + vbYesNo, "To Save")
        sql = "insert into bus_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + Text2.Text + "','" + Combo2.Text + "'," + Text3.Text + ",'" + Combo3.Text + "')"
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
Text2.Text = " "
Text2.Enabled = False
Combo2.ListIndex = 0
Text3.Text = " "
Combo3.ListIndex = 0
add.Enabled = False
save.Enabled = False
update.Enabled = False
search.Enabled = False
End If
End Sub

'To Generate Report

Private Sub report_Click()
    rptBtn_click = True
    Load report1
    report1.Show
End Sub
'To Search
Private Sub search_Click()
conn
Text2.Enabled = True
Text2.Text = " "
Combo2.ListIndex = 0
Text3.Text = " "
Combo3.ListIndex = 0
z = MsgBox("Do You Want To Search Bus Details", vbQuestion + vbYesNo, "To Search Bus Details")
If z = vbYes Then
add.Enabled = False
update.Enabled = True
search.Enabled = False
Combo1.Locked = True
Text1.Locked = True
sql = "select * from bus_mst where  b_id='" & Text1.Text & "'"
Set r = c.Execute(sql)
If r.EOF = 1 Then
   MsgBox "No Matching Details Found. Please Check The Route Id & Bus Id.", vbOKOnly, "To Search Bus Details"
   Combo1.Locked = False
   Combo1.ListIndex = 0
   Text1.Locked = False
   Text1.Text = ""
   Text2.Enabled = False
   Exit Sub
End If
MsgBox "Bus Details Found", vbOKOnly, "To Search Bus Details"
Text2.Text = r.Fields(2)
Combo2.Text = r.Fields(3)
Text3.Text = r.Fields(4)
Combo3.Text = r.Fields(5)
End If
End Sub

'To Update
Private Sub update_Click()
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Route Id To Proceed.", vbOKOnly, "Route Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Bus Id To Find Bus Details.", vbOKOnly, "Bus Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Bus Name To Proceed.", vbOKOnly, "Bus Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Enter Type Of Bus To Proceed.", vbOKOnly, "Bus Type Cannot Be Empty")
        Combo2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Total Number Of Seat To Proceed.", vbOKOnly, "Total Number Of Seats Cannot Be Empty")
        Text3.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo3.SetFocus
    Else

    conn
    p = MsgBox(" Do You Want To Update Bus Details.", vbQuestion + vbYesNo, "To Update")
        sql = "update bus_mst set b_nm='" & Text2.Text & "',T_o_b='" & Combo2.Text & "',T_n_o_s=" & Text3.Text & ",status='" & Combo3.Text & "' where (R_id='" & Combo1.Text & "')and (B_id='" & Text1.Text & "')"
    If p = vbYes Then
        Set r = c.Execute(sql)
        MsgBox "Record Updated", vbOKOnly, "To Update"
        Else
        Exit Sub
    End If
    Adodc1.Refresh
    Combo1.Locked = False
    Combo1.ListIndex = 0
    Text1.Text = " "
    Text2.Text = " "
    Combo2.ListIndex = 0
    Text3.Text = ""
    Combo3.ListIndex = 0
    update.Enabled = False
    End If
End Sub

' To Clear

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
Combo2.ListIndex = 0
Text3.Text = ""
Combo3.ListIndex = 0
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

'To Exit
Private Sub exit_Click()
u = MsgBox("Do You Want To Save Changes Before Exiting", vbQuestion + vbYesNoCancel, "To Exit")
If u = vbYes Then
If Combo1.ListIndex = 0 Then
        u = MsgBox("Please Select A Route Id To Proceed.", vbOKOnly, "Route Id Cannot Be Empty")
        Combo1.SetFocus
    ElseIf Text1.Text = "" Then
        u = MsgBox("Please Click Add New Or Enter Bus Id To Find Bus Details.", vbOKOnly, "Bus Id Cannot Be Empty")
        Text1.SetFocus
   ElseIf Text2.Text = "" Then
        u = MsgBox("Please Enter Bus Name To Proceed.", vbOKOnly, "Bus Name Cannot Be Empty")
        Text2.SetFocus
   ElseIf Combo2.ListIndex = 0 Then
        u = MsgBox("Please Enter Type Of Bus To Proceed.", vbOKOnly, "Bus Type Cannot Be Empty")
        Combo2.SetFocus
   ElseIf Text3.Text = "" Then
        u = MsgBox("Please Enter Total Number Of Seat To Proceed.", vbOKOnly, "Total Number Of Seats Cannot Be Empty")
        Text3.SetFocus
   ElseIf Combo3.ListIndex = 0 Then
        u = MsgBox("Please Select A Valid Status.", vbOKOnly, "Select Status")
        Combo3.SetFocus
    Else
conn
sql = "insert into bus_mst values ('" + Combo1.Text + "','" + Text1.Text + "','" + Text2.Text + "','" + Combo2.Text + "'," + Text3.Text + ",'" + Combo3.Text + "')"
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



Private Sub Form_Load()
conn
Combo1.AddItem ("Select Route Id")
Combo1.ListIndex = 0
Combo2.AddItem ("Select Bus Type")
Combo2.AddItem ("AC")
Combo2.AddItem ("NON-AC")
Combo2.ListIndex = 0
Combo3.AddItem ("Select Status")
Combo3.AddItem ("TRUE")
Combo3.AddItem ("FALSE")
Combo3.ListIndex = 0
Unload booking1
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
sql = "select R_id From route_mst"
Set r = c.Execute(sql)
While r.EOF <> True
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
add.Enabled = False
save.Enabled = False
search.Enabled = False
update.Enabled = False
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

Public Sub gen()
sql = "select max (to_number(SUBSTR(B_id,4,length(B_id)))) from bus_mst"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
Text1.Text = "B" & "00" & 1
Else
Text1.Text = "B" & "00" & r.Fields(0) + 1
End If
a = Text1.Text
If (a = "B" & "001" & 0) Then
sql = "select max (to_number(SUBSTR(B_id,3,length(B_id)))) from bus_mst"
Set r = c.Execute(sql)
Text1.Text = "B" & "0" & r.Fields(0) + 1
End If
End Sub

' For only character
Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 64 And KeyAscii < 91 Or KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 20 Or KeyAscii = 8 Or KeyAscii = 32 Then
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

' For Upper Case
Private Sub Text1_LostFocus()
If Combo1.ListIndex = 0 Then
u = MsgBox("Please Select A Route Id To Proceed.", vbOKOnly, "Route Id Cannot Be Empty")
Combo1.SetFocus
End If
Text1.Text = UCase(Text1.Text)
End Sub


