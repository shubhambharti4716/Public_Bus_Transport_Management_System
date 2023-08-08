VERSION 5.00
Begin VB.Form report1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14640
   Icon            =   "report1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   14640
   WindowState     =   2  'Maximized
   Begin VB.CommandButton getreport 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Get Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5400
      Width           =   5775
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3720
      Width           =   5775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose To Generate Report On The Basis Of :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6000
      TabIndex        =   5
      Top             =   4800
      Width           =   5775
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select One Of The Following   :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6000
      TabIndex        =   4
      Top             =   3120
      Width           =   4770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT GENERATING FORM"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   7425
   End
End
Attribute VB_Name = "report1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
conn
Combo1.AddItem ("Select One Of The Following")
Combo1.AddItem ("Route Details")
Combo1.AddItem ("Stoppage Details")
Combo1.AddItem ("Bus Details")
Combo1.AddItem ("Driver Details")
Combo1.AddItem ("Conductor Details")
Combo1.AddItem ("Trip Details")
Combo1.AddItem ("Customer Details")
Combo1.AddItem ("Booking Details")
Combo1.AddItem ("Bill Details")
Combo1.ListIndex = 0
Combo2.AddItem ("Choose To Generate Report On The Basis Of")
Combo2.ListIndex = 0
Unload booking1
Unload bus1
Unload conductor1
Unload customer1
Unload driver1
Unload route1
Unload stoppage1
Unload bill1
Unload trip1
HOMEPAGE.Picture2.Visible = True
HOMEPAGE.Picture3.Visible = False
check_for_activeform
formopen = 1
Set dataenviroment = Nothing
End Sub
'To Unload
Private Sub Form_QueryUnload(Cancel As Integer, unloadmode As Integer)
   If mdiBtn_click = True Then
      Exit Sub
      HOMEPAGE.Picture3.Visible = True
   End If

If MsgBox("Are you sure you want to exit ?", vbQuestion + vbYesNo, "To Exit") = vbNo Then
Cancel = True
Exit Sub
End If
HOMEPAGE.Picture3.Visible = True
End Sub



Private Sub Combo1_click()
Combo2.clear
If Combo1.ListIndex = 0 Then
    Combo2.AddItem ("Choose To Generate Report On The Basis Of")
Else
If Combo1.ListIndex = 1 Then
        'Route
        Combo2.AddItem ("Choose To Generate Report On The Basis Of")
        Combo2.AddItem ("On The Basis Of Particular Route Id")
        Combo2.AddItem ("On The Basis Of All Route Id")
Else
If Combo1.ListIndex = 2 Then
        'Stoppage
        Combo2.AddItem ("Choose To Generate Report On The Basis Of")
        Combo2.AddItem ("On The Basis Of Particular Route Id")
        Combo2.AddItem ("On The Basis Of All Route Id")
        Combo2.AddItem ("On The Basis Of Particular Stoppage Id")
        Combo2.AddItem ("On The Basis Of All Stoppage Id")
Else
If Combo1.ListIndex = 3 Then
        'Bus
        Combo2.AddItem ("Choose To Generate Report On The Basis Of")
        Combo2.AddItem ("On The Basis Of Particular Route Id")
        Combo2.AddItem ("On The Basis Of All Route Id")
        Combo2.AddItem ("On The Basis Of Particular Bus Id")
        Combo2.AddItem ("On The Basis Of All Bus Id")
Else
If Combo1.ListIndex = 4 Then
        'Driver
        Combo2.AddItem ("Choose To Generate Report On The Basis Of")
        Combo2.AddItem ("On The Basis Of Particular Bus Id")
        Combo2.AddItem ("On The Basis Of All Bus Id")
        Combo2.AddItem ("On The Basis Of Particular Driver Id")
        Combo2.AddItem ("On The Basis Of All Driver Id")
Else
If Combo1.ListIndex = 5 Then
        'Conductor
        Combo2.AddItem ("Choose To Generate Report On The Basis Of")
        Combo2.AddItem ("On The Basis Of Particular Bus Id")
        Combo2.AddItem ("On The Basis Of All Bus Id")
        Combo2.AddItem ("On The Basis Of Particular Conductor Id")
        Combo2.AddItem ("On The Basis Of All Conductor Id")
Else
If Combo1.ListIndex = 6 Then
        'Trip
        Combo2.AddItem ("Choose To Generate Report On The Basis Of")
        Combo2.AddItem ("On The Basis Of Particular Bus Id")
        Combo2.AddItem ("On The Basis Of All Bus Id")
        Combo2.AddItem ("On The Basis Of Particular Trip Id")
        Combo2.AddItem ("On The Basis Of All Trip Id")
Else
If Combo1.ListIndex = 7 Then
        'Customer
        Combo2.AddItem ("Choose To Generate Report On The Basis Of")
        Combo2.AddItem ("On The Basis Of Particular Customer Id")
        Combo2.AddItem ("On The Basis Of All Customer Id")
        Combo2.AddItem ("On The Basis Of Particular Phone Number")
Else
If Combo1.ListIndex = 8 Then
        'Booking
        Combo2.AddItem ("Choose To Generate Report On The Basis Of")
        Combo2.AddItem ("On The Basis Of Particular Bus Id")
        Combo2.AddItem ("On The Basis Of All Bus Id")
        Combo2.AddItem ("On The Basis Of Particular Customer Id")
        Combo2.AddItem ("On The Basis Of All Customer Id")
        Combo2.AddItem ("On The Basis Of Particular Booking Id")
        Combo2.AddItem ("On The Basis Of All Booking Id")
        Combo2.AddItem ("On The Basis Of Particular Date For Booking")
        Combo2.AddItem ("On The Basis Of All Date For Booking")
Else
If Combo1.ListIndex = 9 Then
        'Bill
        Combo2.AddItem ("Choose To Generate Report On The Basis Of")
        Combo2.AddItem ("On The Basis Of Particular Booking Id")
        Combo2.AddItem ("On The Basis Of All Booking Id")
        Combo2.AddItem ("On The Basis Of Particular Date For Booking")
        Combo2.AddItem ("On The Basis Of Particular Bill Id")
        Combo2.AddItem ("On The Basis Of All Bill Id")
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub getreport_Click()
Set DataEnvironment1 = Nothing
Set DataEnvironment2 = Nothing
Set DataEnvironment3 = Nothing
Set DataEnvironment4 = Nothing
Set DataEnvironment5 = Nothing
Set DataEnvironment6 = Nothing
Set DataEnvironment7 = Nothing
Set DataEnvironment8 = Nothing
Dim r As String
Dim mydate As Date
If DataEnvironment1.Connection1.State = 1 Then
    DataEnvironment1.Connection1.Close
Else
If DataEnvironment1.Connection1.State = 0 Then
    DataEnvironment1.Connection1.Open
    
If DataEnvironment2.Connection1.State = 1 Then
    DataEnvironment2.Connection1.Close
Else
If DataEnvironment2.Connection1.State = 0 Then
    DataEnvironment2.Connection1.Open
    
If DataEnvironment3.Connection1.State = 1 Then
    DataEnvironment3.Connection1.Close
Else
If DataEnvironment3.Connection1.State = 0 Then
    DataEnvironment3.Connection1.Open
    
If DataEnvironment4.Connection1.State = 1 Then
    DataEnvironment4.Connection1.Close
Else
If DataEnvironment4.Connection1.State = 0 Then
    DataEnvironment4.Connection1.Open
    
If DataEnvironment5.Connection1.State = 1 Then
    DataEnvironment5.Connection1.Close
Else
If DataEnvironment5.Connection1.State = 0 Then
    DataEnvironment5.Connection1.Open
    
If DataEnvironment6.Connection1.State = 1 Then
    DataEnvironment6.Connection1.Close
Else
If DataEnvironment6.Connection1.State = 0 Then
    DataEnvironment6.Connection1.Open
    
If DataEnvironment7.Connection1.State = 1 Then
    DataEnvironment7.Connection1.Close
Else
If DataEnvironment7.Connection1.State = 0 Then
    DataEnvironment7.Connection1.Open
    
If DataEnvironment8.Connection1.State = 1 Then
    DataEnvironment8.Connection1.Close
Else
If DataEnvironment8.Connection1.State = 0 Then
    DataEnvironment8.Connection1.Open
    'For Route
    If Combo1.ListIndex = 1 Then
        If Combo2.ListIndex = 1 Then
            r = InputBox("Enter the Route Id", "Report")
            r = UCase(r)
            DataEnvironment1.route (r)
            DataReport1.Show
        Else
        If Combo2.ListIndex = 2 Then
            DataReport2.Show
        End If
        End If
    Else
    'For Stoppage
    If Combo1.ListIndex = 2 Then
        If Combo2.ListIndex = 1 Then
            r = InputBox("Enter the Route Id", "Report")
            r = UCase(r)
            DataEnvironment1.stoppage (r)
            DataReport3.Show
        Else
        If Combo2.ListIndex = 2 Then
            DataEnvironment1.stoppage (r)
            DataReport4.Show
        Else
        If Combo2.ListIndex = 3 Then
            r = InputBox("Enter the Stoppage Id", "Report")
            r = UCase(r)
            DataEnvironment3.stoppage (r)
            DataReport5.Show
        Else
        If Combo2.ListIndex = 4 Then
            DataReport6.Show
        End If
        End If
        End If
        End If
    Else
    'For Bus
    If Combo1.ListIndex = 3 Then
        If Combo2.ListIndex = 1 Then
        r = InputBox("Enter the Route Id", "Report")
            r = UCase(r)
            DataEnvironment1.bus (r)
            DataReport7.Show
        Else
        If Combo2.ListIndex = 2 Then
            DataReport8.Show
        Else
        If Combo2.ListIndex = 3 Then
            r = InputBox("Enter the Bus Id", "Report")
            r = UCase(r)
            DataEnvironment3.bus (r)
            DataReport9.Show
        Else
        If Combo2.ListIndex = 4 Then
            DataReport10.Show
        End If
        End If
        End If
        End If
    Else
    'For Driver
    If Combo1.ListIndex = 4 Then
        If Combo2.ListIndex = 1 Then
        r = InputBox("Enter the Bus Id", "Report")
            r = UCase(r)
            DataEnvironment1.driver (r)
            DataReport11.Show
        Else
        If Combo2.ListIndex = 2 Then
            DataReport12.Show
        Else
        If Combo2.ListIndex = 3 Then
        r = InputBox("Enter the Driver Id", "Report")
            r = UCase(r)
            DataEnvironment3.driver (r)
            DataReport13.Show
        Else
        If Combo2.ListIndex = 4 Then
            DataReport14.Show
        End If
        End If
        End If
        End If
    Else
    'For Conductor
    If Combo1.ListIndex = 5 Then
        If Combo2.ListIndex = 1 Then
        r = InputBox("Enter the Bus Id", "Report")
            r = UCase(r)
            DataEnvironment1.conductor (r)
            DataReport15.Show
        Else
        If Combo2.ListIndex = 2 Then
            DataReport16.Show
        Else
        If Combo2.ListIndex = 3 Then
        r = InputBox("Enter the Conductor Id", "Report")
            r = UCase(r)
            DataEnvironment3.conductor (r)
            DataReport17.Show
        Else
        If Combo2.ListIndex = 4 Then
            DataReport18.Show
        End If
        End If
        End If
        End If
    Else
    'For Trip
    If Combo1.ListIndex = 6 Then
        If Combo2.ListIndex = 1 Then
        r = InputBox("Enter the Bus Id", "Report")
            r = UCase(r)
            DataEnvironment1.trip (r)
            DataReport19.Show
        Else
        If Combo2.ListIndex = 2 Then
            DataReport20.Show
        Else
        If Combo2.ListIndex = 3 Then
        r = InputBox("Enter the Trip Id", "Report")
            r = UCase(r)
            DataEnvironment3.trip (r)
            DataReport21.Show
        Else
        If Combo2.ListIndex = 4 Then
            DataReport22.Show
        End If
        End If
        End If
        End If
    Else
    'For Customer
    If Combo1.ListIndex = 7 Then
        If Combo2.ListIndex = 1 Then
        r = InputBox("Enter the Customer Id", "Report")
            r = UCase(r)
            DataEnvironment1.customer (r)
            DataReport23.Show
        Else
        If Combo2.ListIndex = 2 Then
            DataReport24.Show
        Else
        If Combo2.ListIndex = 3 Then
        r = InputBox("Enter the Customer Phone Number", "Report")
            r = UCase(r)
            DataEnvironment3.customer (r)
            DataReport25.Show
        End If
        End If
        End If
    Else
    'For Booking
    If Combo1.ListIndex = 8 Then
        If Combo2.ListIndex = 1 Then
        r = InputBox("Enter the Bus Id", "Report")
            r = UCase(r)
            DataEnvironment1.booking (r)
            DataReport26.Show
        Else
        If Combo2.ListIndex = 2 Then
            DataReport27.Show
        Else
        If Combo2.ListIndex = 3 Then
        r = InputBox("Enter the Customer Id", "Report")
            r = UCase(r)
            DataEnvironment3.booking (r)
            DataReport28.Show
        Else
        If Combo2.ListIndex = 4 Then
            DataReport29.Show
        Else
        If Combo2.ListIndex = 5 Then
            DataReport30.Show
        Else
        If Combo2.ListIndex = 6 Then
            DataReport31.Show
        Else
        If Combo2.ListIndex = 7 Then
       ' r = mydate("Select Particular Date For Booking Id", "Report")
        'r = DateSerial(r)
            DataReport32.Show
        Else
        If Combo2.ListIndex = 8 Then
            DataReport33.Show
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
    Else
    'For Bill
    If Combo1.ListIndex = 9 Then
        If Combo2.ListIndex = 1 Then
            r = InputBox("Enter the Booking Id", "Report")
            r = UCase(r)
            DataEnvironment1.bill (r)
            DataReport34.Show
        Else
        If Combo2.ListIndex = 2 Then
            DataReport35.Show
        Else
        If Combo2.ListIndex = 3 Then
            DataReport36.Show
        Else
        If Combo2.ListIndex = 4 Then
        r = InputBox("Enter the Bill Id", "Report")
            r = UCase(r)
            DataEnvironment4.bill (r)
            DataReport37.Show
        Else
        If Combo2.ListIndex = 5 Then
            DataReport38.Show
        End If
        End If
        End If
        End If
        End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
Combo1.ListIndex = 0
End Sub


