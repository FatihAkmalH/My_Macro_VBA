Sub RefreshAndCleanPivotTablesWithMultipleConnections()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim pivotSheetNames As Variant
    Dim targetSourceData As String
    Dim i As Long
    Dim missingSheets As String
    Dim processedSheets As String
    
    ' Nama sheet yang digunakan (bisa lebih dari satu)
    pivotSheetNames = Array("PIVOT", "ANALISIS")
    
    ' Sumber data baru untuk PivotTable
    targetSourceData = "DATA!Table_DataBaru"
    
    ' Refresh semua koneksi data
    ActiveWorkbook.RefreshAll
    
    ' Loop semua sheet dalam daftar
    For i = LBound(pivotSheetNames) To UBound(pivotSheetNames)
        
        On Error Resume Next
        Set ws = ActiveWorkbook.Sheets(pivotSheetNames(i))
        On Error GoTo 0
        
        If ws Is Nothing Then
            missingSheets = missingSheets & vbCrLf & "- " & pivotSheetNames(i)
        Else
            ' Loop semua PivotTable di sheet
            For Each pt In ws.PivotTables
                If pt.SourceData = targetSourceData Then
                    pt.RefreshTable
                    
                    ' Bersihkan field pivot
                    For Each pf In pt.PivotFields
                        If pf.Orientation = xlRowField Or pf.Orientation = xlColumnField Or pf.Orientation = xlPageField Then
                            On Error Resume Next
                            
                            ' Tampilkan semua item
                            For Each pi In pf.PivotItems
                                pi.Visible = True
                            Next pi
                            
                            ' Sembunyikan item tertentu
                            If pf.PivotItems("(blank)").Visible = True Then pf.PivotItems("(blank)").Visible = False
                            If pf.PivotItems("").Visible = True Then pf.PivotItems("").Visible = False
                            If pf.PivotItems("0").Visible = True Then pf.PivotItems("0").Visible = False
                            If pf.PivotItems(" ").Visible = True Then pf.PivotItems(" ").Visible = False
                            
                            On Error GoTo 0
                        End If
                    Next pf
                End If
            Next pt
            
            processedSheets = processedSheets & vbCrLf & "- " & pivotSheetNames(i)
        End If
    Next i
    
    ' Tampilkan ringkasan hasil
    If processedSheets <> "" Then
        MsgBox "PivotTable berhasil di-refresh dan difilter pada sheet berikut:" & vbCrLf & processedSheets, vbInformation
    End If
    
    If missingSheets <> "" Then
        MsgBox "Sheet berikut tidak ditemukan:" & vbCrLf & missingSheets, vbExclamation
    End If
End Sub