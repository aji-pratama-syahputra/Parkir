Attribute VB_Name = "Module1"

Option Explicit

'No. Plat
Public cNoPlat As String

'Koneksi ADO
Public oConn As New ADODB.Connection

'String SQL
Public S As String

'Temp Int
Public i As Integer

'User Name
Public strUser As String

'User Type
Public strUserType As String

'Laporan
Public BTgl As Boolean
Public Tgl1 As Date
Public Tgl2 As Date
Public BJam As Boolean
Public Jam1 As String
Public Jam2 As String
Public BOp As Boolean
Public Op1 As String
Public Op2 As String

Sub Main()
    
    'Buka koneksi ke database
    oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                             App.Path & "\SecurePark.mdb;Persist Security Info=False"
    oConn.Open
    
    'Show Form Login
    frmLogin.Show
    
End Sub

Public Sub CloseRS(pRS As ADODB.Recordset)
    pRS.Close
    Set pRS = Nothing
End Sub

Public Function FormatTgl(pDate As Date) As String
    FormatTgl = "#" & Format(pDate, "YYYY-MM-DD") & "#"
End Function

Public Sub GotFocus(pOTextBox As TextBox)
    With pOTextBox
        .Text = Format(.Text, "0")
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Public Sub LostFocus(pOTextBox As TextBox)
    With pOTextBox
        If Trim(.Text) = "" Then .Text = "0"
        .Text = Format(.Text, "#,##0")
    End With
End Sub
