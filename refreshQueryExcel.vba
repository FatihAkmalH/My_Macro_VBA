Sub RefreshQueryTable()
    Dim ws As Worksheet
    Dim tbl As ListObject

    ' Tentukan sheet QUERY
    Set ws = ThisWorkbook.Sheets("QUERY")
    
    ' Pastikan query table bernama "tbl_query" ada di sheet QUERY
    On Error Resume Next
    Set tbl = ws.ListObjects("tbl_query") 'ganti nama worksheet
    On Error GoTo 0
    
    ' Jika query table ditemukan, lakukan refresh
    If Not tbl Is Nothing Then
        tbl.QueryTable.Refresh BackgroundQuery:=False
        MsgBox "Query 'tbl_query' berhasil diperbarui!", vbInformation, "Refresh Berhasil"
    Else
        MsgBox "Query table 'tbl_query' tidak ditemukan di sheet 'QUERY'.", vbExclamation, "Kesalahan"
    End If
End Sub