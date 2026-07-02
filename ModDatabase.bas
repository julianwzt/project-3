Attribute VB_Name = "ModDatabase"
Public conn As New ADODB.Connection
Public rs As New ADODB.Recordset

Public Sub BukaKoneksi()
    On Error GoTo ErrorHandler
    
    If conn.State = 1 Then conn.Close
    
    conn.Open "DSN=db_mahasiswa;UID=postgres;PWD=20060721;"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Koneksi ke PostgreSQL Gagal: " & Err.Description, vbCritical, "Error Koneksi"
End Sub

Public Sub TutupKoneksi()
    If conn.State = 1 Then conn.Close
End Sub

Public Sub TampilData(Grid As DataGrid)
    On Error GoTo ErrorHandler

    Call BukaKoneksi

    Set rs = New ADODB.Recordset

    rs.CursorLocation = adUseClient

    Dim query As String
    query = "SELECT m.id, m.nama, m.umur, m.nim, m.tgl_lahir, m.alamat, j.nama_jurusan " & _
            "FROM Mahasiswa m " & _
            "LEFT JOIN Jurusan j ON m.id_jurusan = j.id_jurusan ORDER BY m.id DESC"

    rs.Open query, conn, adOpenStatic, adLockReadOnly

    Set Grid.DataSource = rs
    
    Exit Sub
ErrorHandler:
    MsgBox "Gagal memuat data: " & Err.Description, vbCritical, "Error"
End Sub

