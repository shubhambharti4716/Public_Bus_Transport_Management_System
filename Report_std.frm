VERSION 5.00
Begin VB.Form Report_std 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4500
   Icon            =   "Report_std.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ROUTE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   29
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton command1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Report Generation"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF ROUTE ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF ROUTE NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   3495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "STOPPAGE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   25
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Report Generation"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF STOPPAGE ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   28
         Top             =   840
         Width           =   3135
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF ROUTE ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   3855
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF STOPPAGE NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   26
         Top             =   1440
         Width           =   3495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BUS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   22
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Report Generation"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF BUS NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   32
         Top             =   1440
         Width           =   3015
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF BUS ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF ROUTE ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DRIVER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   18
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Report Generation"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF DRIVER NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   21
         Top             =   1440
         Width           =   3375
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF BUS ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   2895
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS DRIVER ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   19
         Top             =   840
         Width           =   3375
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CONDUCTOR"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Report Generation"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS CONDUCTOR ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   3735
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF BUS ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   3495
      End
      Begin VB.OptionButton Option14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF CONDUCTOR NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   3975
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CUSTOMER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Report Generation"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton Option15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF CUSTOMER ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   3615
      End
      Begin VB.OptionButton Option16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF CUSTOMER NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   3735
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TRIP"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Report Generation"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.OptionButton Option17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF BUS ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   3135
      End
      Begin VB.OptionButton Option18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF TRIP ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   3015
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BOOKING"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Report Generation"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.OptionButton Option20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS CUSTOMER  ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   3255
      End
      Begin VB.OptionButton Option19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF BUS ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   2895
      End
      Begin VB.OptionButton Option21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF BOOKING ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   3495
      End
      Begin VB.OptionButton Option22 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF DATE FOR BOOKING"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Width           =   3855
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TICKET"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command9 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Report Generation"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option24 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF TICKET ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   3495
      End
      Begin VB.OptionButton Option23 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON THE BASIS OF BUS ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Report_std"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
Dim r As String
If Option1.Value = True Then
r = InputBox("Enter the Route Id", "Report")
r = UCase(r)
   If DataEnvironment1.Connection1.State = 1 Then DataEnvironment1.Connection1.Close
   DataEnvironment1.route (r)
   DataReport1.Show
   Unload Me
Else
If Option2.Value = True Then
DataReport2.Show
Unload Me
End If
End If
End Sub

Private Sub Command2_Click()
If Option3.Value = True Then
DataReport3.Show
Unload Me
Else
If Option4.Value = True Then
DataReport4.Show
Unload Me
Else
If Option5.Value = True Then
DataReport5.Show
Unload Me
End If
End If
End If
End Sub

Private Sub Command3_Click()
If Option6.Value = True Then
DataReport6.Show
Else
If Option7.Value = True Then
DataReport7.Show
Else
If Option8.Value = True Then
DataReport8.Show
End If
End If
End If
End Sub

Private Sub Command4_Click()
If Option9.Value = True Then
DataReport9.Show
Else
If Option10.Value = True Then
DataReport10.Show
Else
If Option11.Value = True Then
DataReport11.Show
End If
End If
End If
End Sub
Private Sub Command5_Click()
If Option12.Value = True Then
DataReport12.Show
Else
If Option13.Value = True Then
DataReport13.Show
Else
If Option14.Value = True Then
DataReport14.Show
End If
End If
End If
End Sub

Private Sub Command6_Click()
If Option15.Value = True Then
DataReport15.Show
Else
If Option16.Value = True Then
DataReport16.Show
End If
End If
End Sub

Private Sub Command7_Click()
If Option17.Value = True Then
DataReport17.Show
Else
If Option18.Value = True Then
DataReport18.Show
End If
End If
End Sub

Private Sub Command8_Click()
If Option19.Value = True Then
DataReport19.Show
Else
If Option20.Value = True Then
DataReport20.Show
Else
If Option21.Value = True Then
DataReport21.Show
Else
If Option22.Value = True Then
DataReport22.Show
End If
End If
End If
End If
End Sub

Private Sub Command9_Click()
If Option23.Value = True Then
DataReport23.Show
Else
If Option24.Value = True Then
DataReport24.Show
End If
End If
End Sub

