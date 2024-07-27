Attribute VB_Name = "Module1"
Public Koneksi As New ADODB.Connection
Public RSUsers As New ADODB.Recordset
Public RSPendaftaran As New ADODB.Recordset


Public Sub BukaDB()
Set Koneksi = New ADODB.Connection
Set RSUsers = New ADODB.Recordset
Set RSPendaftaran = New ADODB.Recordset
    
Koneksi.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\KlinikDB.mdb"
End Sub

