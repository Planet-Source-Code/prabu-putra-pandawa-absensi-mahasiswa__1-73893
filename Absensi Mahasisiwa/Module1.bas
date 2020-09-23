Attribute VB_Name = "Module1"
'-----------------Koneksi DataBase
Public XKoneksi As New ADODB.Connection
Public RsAbsensi As New ADODB.Recordset
Public RsDosen As New ADODB.Recordset
Public RsMahasiswa As New ADODB.Recordset
Public RsMatakuliah As New ADODB.Recordset
Public crreport As New ADODB.Recordset

'-----------------Koneksi ado
Sub koneksi()
With XKoneksi
    .CursorLocation = adUseClient
    .Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;PERSIST SECURITY INFO=FALSE;DATA SOURCE=" & App.Path & "\siswa.mdb"
End With
End Sub

'-----------------Keluar aplikasi
Sub keluar()
    XKoneksi.Close
End Sub
