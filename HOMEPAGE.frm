VERSION 5.00
Begin VB.MDIForm HOMEPAGE 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Public Bus Transport Management System"
   ClientHeight    =   7950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20250
   Icon            =   "HOMEPAGE.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7950
      Left            =   17310
      Picture         =   "HOMEPAGE.frx":26E8E
      ScaleHeight     =   530
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   5
      Top             =   0
      Width           =   2940
      Begin VB.CommandButton exit 
         BackColor       =   &H0080FF80&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Open Home Page"
         Top             =   9360
         Width           =   2055
      End
      Begin VB.CommandButton HOM_CO7 
         BackColor       =   &H0080FF80&
         Caption         =   "Trip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Open Trip Menu"
         Top             =   6360
         Width           =   2055
      End
      Begin VB.CommandButton HOM_CO2 
         BackColor       =   &H0080FF80&
         Caption         =   "Stoppage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Open Stoppage Menu"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton HOM_CO1 
         BackColor       =   &H0080FF80&
         Caption         =   "Route"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Open Route  Menu"
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton HOM_CO3 
         BackColor       =   &H0080FF80&
         Caption         =   "Bus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Open Bus Menu"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton HOM_CO4 
         BackColor       =   &H0080FF80&
         Caption         =   "Driver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Open Driver menu"
         Top             =   5160
         Width           =   2055
      End
      Begin VB.CommandButton HOM_CO5 
         BackColor       =   &H0080FF80&
         Caption         =   "Conductor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Open Conductor Menu"
         Top             =   5760
         Width           =   2055
      End
      Begin VB.CommandButton HOM_CO6 
         BackColor       =   &H0080FF80&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Open Customer Menu"
         Top             =   6960
         Width           =   2055
      End
      Begin VB.CommandButton HOM_CO8 
         BackColor       =   &H0080FF80&
         Caption         =   "Booking"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Open Booking Menu"
         Top             =   7560
         Width           =   2055
      End
      Begin VB.CommandButton HOM_CO9 
         BackColor       =   &H0080FF80&
         Caption         =   "Bill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Open Ticket Menu"
         Top             =   8160
         Width           =   2055
      End
      Begin VB.CommandButton HOM_CO10 
         BackColor       =   &H0080FF80&
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Open Home Page"
         Top             =   8760
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   3  'Align Left
      Height          =   7950
      Left            =   0
      Picture         =   "HOMEPAGE.frx":2C072
      ScaleHeight     =   7890
      ScaleWidth      =   17475
      TabIndex        =   0
      Top             =   0
      Width           =   17535
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   20250
         TabIndex        =   1
         Top             =   0
         Width           =   20250
         Begin VB.Timer Timer1 
            Interval        =   100
            Left            =   240
            Top             =   0
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   " HELLO AND WELCOME  TO PUBLIC TRANSPORT MANAGEMENT SYSTEM"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   2760
            TabIndex        =   4
            Top             =   0
            Width           =   12405
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   15240
            TabIndex        =   3
            Top             =   0
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   2295
         End
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   14760
      Top             =   0
   End
   Begin VB.Menu route 
      Caption         =   "Route"
   End
   Begin VB.Menu stoppage 
      Caption         =   "Stoppage"
   End
   Begin VB.Menu bus 
      Caption         =   "Bus"
   End
   Begin VB.Menu driver 
      Caption         =   "Driver"
   End
   Begin VB.Menu conductor 
      Caption         =   "Conductor"
   End
   Begin VB.Menu trip 
      Caption         =   "Trip"
   End
   Begin VB.Menu customer 
      Caption         =   "Customer"
   End
   Begin VB.Menu booking 
      Caption         =   "Booking"
   End
   Begin VB.Menu bill 
      Caption         =   "Bill"
   End
   Begin VB.Menu report 
      Caption         =   "Report"
   End
End
Attribute VB_Name = "HOMEPAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bill_Click()
Load bill1
bill1.Show
End Sub

Private Sub booking_Click()
Load booking1
booking1.Show
End Sub

'to show bus form
Private Sub bus_Click()
Load bus1
bus1.Show
End Sub

'Private Sub exit_Click()
 '  p = MsgBox("Do you want to exit", vbYesNo, "On Exit")
  '   If p = yes Then
   '     Unload Me
    '  Else
     '   Exit Sub
   ' End If
'End Sub


Private Sub conductor_Click()
Load conductor1
conductor1.Show
End Sub

Private Sub customer_Click()
Load customer1
customer1.Show
End Sub

Private Sub driver_Click()
Load driver1
driver1.Show
End Sub

Private Sub HOM_CO1_Click()
Load route1
route1.Show
End Sub


Private Sub HOM_CO10_Click()
If HOMEPAGE.Picture3.Visible = True Then
Y = MsgBox("You Are On The HomePage", vbOKOnly, "Homepage")
Else
Y = MsgBox("Do You Want To Close This Window", vbQuestion + vbYesNo, "To Exit")
If Y = vbYes Then
Unload booking1
Unload bus1
Unload conductor1
Unload customer1
Unload driver1
Unload report1
Unload route1
Unload stoppage1
Unload bill1
Unload trip1
HOMEPAGE.Picture3.Visible = True
End If
End If
End Sub

Private Sub HOM_CO2_Click()
Y = MsgBox("Do You Want To Close This Window", vbQuestion + vbYesNo, "To Exit")
If Y = vbYes Then
Load stoppage1
stoppage1.Show
End If
End Sub

Private Sub HOM_CO3_Click()
Y = MsgBox("Do You Want To Close This Window", vbQuestion + vbYesNo, "To Exit")
If Y = vbYes Then
Load bus1
bus1.Show
End If
End Sub

Private Sub HOM_CO4_Click()
Y = MsgBox("Do You Want To Close This Window", vbQuestion + vbYesNo, "To Exit")
If Y = vbYes Then
Load driver1
driver1.Show
End If
End Sub

Private Sub HOM_CO5_Click()
Y = MsgBox("Do You Want To Close This Window", vbQuestion + vbYesNo, "To Exit")
If Y = vbYes Then
Load conductor1
conductor1.Show
End If
End Sub

Private Sub HOM_CO6_Click()
Y = MsgBox("Do You Want To Close This Window", vbQuestion + vbYesNo, "To Exit")
If Y = vbYes Then
Load customer1
customer1.Show
End If
End Sub

Private Sub HOM_CO7_Click()
Y = MsgBox("Do You Want To Close This Window", vbQuestion + vbYesNo, "To Exit")
If Y = vbYes Then
Load trip1
trip1.Show
End If
End Sub

Private Sub HOM_CO8_Click()
Y = MsgBox("Do You Want To Close This Window", vbQuestion + vbYesNo, "To Exit")
If Y = vbYes Then
Load booking1
booking1.Show
End If
End Sub

Private Sub HOM_CO9_Click()
Y = MsgBox("Do You Want To Close This Window", vbQuestion + vbYesNo, "To Exit")
If Y = vbYes Then
Load bill1
bill1.Show
End If
End Sub


Private Sub MDIForm_Load()
Label3.Caption = "      WELCOME TO PUBLIC BUS TRANSPORT MANAGEMENT SYSTEM                    "
Timer1.Enabled = True
Timer1.Interval = 220
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, unloadmode As Integer)
   mdiBtn_click = True
   If MsgBox("Do you want to exit", vbYesNo, "On Exit") = vbNo Then
      Cancel = 1
   End If
End Sub





Private Sub report_Click()
Load report1
report1.Show
End Sub

'goto route form
Private Sub route_Click()
Load route1
route1.Show
End Sub
'goto stoppage form
Private Sub stoppage_Click()
Load stoppage1
stoppage1.Show
End Sub



Private Sub Timer1_Timer()
Dim str As String
str = HOMEPAGE.Label3.Caption
str = Mid$(str, 2, Len(str)) + Left(str, 1)
HOMEPAGE.Label3.Caption = str
End Sub

Private Sub Timer2_Timer()
Label1.Caption = Date
Label2.Caption = Time
End Sub

'To goto Trip
Private Sub trip_Click()
Load trip1
trip1.Show
End Sub
