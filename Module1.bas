Attribute VB_Name = "Module1"
Public c As New ADODB.Connection
Public r As New ADODB.Recordset
Public sql As String
Public gender As String
Public mdiBtn_click As Boolean
Public rptBtn_click As Boolean

'module connection
Public Function conn()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;Password=abc;User ID=sb;Persist Security Info=true"
Set r = New ADODB.Recordset
End Function
Public Sub check_for_activeform()
If formopen = 1 Then
HOMEPAGE.ActiveForm.WindowState = 1
End If
End Sub
