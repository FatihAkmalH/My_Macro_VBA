Sub RawDataMBM()
    Dim srcWb As Workbook, dstWb As Workbook, newWb As Workbook, macroWb As Workbook
    Dim srcWs As Worksheet, dstWs As Worksheet
    Dim srcRange As Range, dstRange As Range
    Dim sheetName As String
    Dim savePath As String
    Dim Weekx As String, Dayx As String, folder_savey As String, file_savey As String
    
    ' Tentukan workbook macro sebagai workbook aktif
    Set macroWb = ThisWorkbook
    
    ' Membuka file yang diperlukan
    Workbooks.Open FileName:="O:\DEVELOPMENT\#aws\Template RAW DATA MBM.xlsm"
    Workbooks.Open FileName:="C:\Export\RAW DATA MBM.xls"
    
    ' Set workbooks
    Set srcWb = Workbooks("RAW DATA MBM.xls")
    Set dstWb = Workbooks("Template RAW DATA MBM.xlsm")
    
    ' Panggil prosedur RenameSheets untuk mengganti nama sheet di RAW DATA MBM
    Call RenameSheets(srcWb)
    
    ' Loop through sheets in source workbook
    For Each srcWs In srcWb.Sheets
        sheetName = srcWs.Name
        
        ' Cek apakah sheet ada di workbook tujuan
        On Error Resume Next
        Set dstWs = dstWb.Sheets(sheetName)
        On Error GoTo 0
        
        If Not dstWs Is Nothing Then
            ' Copy Column D Data
            Set srcRange = srcWs.Range("E4", srcWs.Cells(Rows.Count, "E").End(xlUp))
            If Application.WorksheetFunction.CountA(srcRange) > 0 Then
                Set dstRange = dstWs.Range("E11")
                srcRange.Copy
                dstRange.PasteSpecial Paste:=xlPasteValues
            End If
            
            ' Copy Column C Data
            Set srcRange = srcWs.Range("C4", srcWs.Cells(Rows.Count, "C").End(xlUp))
            If Application.WorksheetFunction.CountA(srcRange) > 0 Then
                Set dstRange = dstWs.Range("B11")
                srcRange.Copy
                dstRange.PasteSpecial Paste:=xlPasteValues
            End If
            
            ' Copy Column B Data
            Set srcRange = srcWs.Range("B4", srcWs.Cells(Rows.Count, "B").End(xlUp))
            If Application.WorksheetFunction.CountA(srcRange) > 0 Then
                Set dstRange = dstWs.Range("C11")
                srcRange.Copy
                dstRange.PasteSpecial Paste:=xlPasteValues
            End If
            
            ' Copy Column A Data with TRIM and LEFT(A,10)
            Set srcRange = srcWs.Range("A4", srcWs.Cells(Rows.Count, "A").End(xlUp))
            If Application.WorksheetFunction.CountA(srcRange) > 0 Then
                Dim cell As Range
                For Each cell In srcRange
                    cell.Value = Trim(Left(cell.Value, 10))
                Next cell
                Set dstRange = dstWs.Range("D11")
                srcRange.Copy
                dstRange.PasteSpecial Paste:=xlPasteValues
            End If
        End If
        
        ' Reset dstWs for next iteration
        Set dstWs = Nothing
    Next srcWs
    
    ' Ambil nilai Weekx dan Dayx dari workbook macro
    Weekx = macroWb.Sheets(1).Cells(10, 5).Value
    Dayx = macroWb.Sheets(1).Cells(8, 6).Value
    
    ' Tentukan folder dan nama file untuk penyimpanan
    folder_savey = "O:\DEVELOPMENT\#HASIL BY MINUTE\PROGRAM WEEK " & Weekx & "\#EXCEL BY MINUTE PER DAY\"
    file_savey = "Raw Data MBM (" & Dayx & ") - National Urban.xlsm"
    savePath = folder_savey & file_savey
    
    ' Simpan file hasil
    dstWb.SaveAs FileName:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    ' Menutup file sumber
    srcWb.Close False
    dstWb.Close False
    
    ' Buka file hasil
    Set newWb = Workbooks.Open(savePath)
    
    ' Cleanup
    Application.CutCopyMode = False
    MsgBox "Data telah berhasil disalin dan disimpan di: " & savePath, vbInformation
End Sub

Sub RenameSheets(wb As Workbook)
    Dim i As Integer
    Dim Done As Boolean
    Dim n As String
    Dim ws As Worksheet
    Dim category As String
    
    i = 1
    
    ' Loop melalui semua sheet
    For Each ws In wb.Sheets
        ' Lewati dua sheet pertama
        If i > 2 Then
            ' Pastikan nama sheet sesuai pola
            If ws.Name Like "Time split by_ 1 min.*" Then
                ' Ambil nilai dari sel A4 dan D4
                n = Trim(ws.Range("A4").Value)
                category = Trim(ws.Range("D4").Value)
                
                ' Periksa nilai di sel D4
                If category = "MDTV" Then
                    ' Ganti nama berdasarkan nilai di A4
                    Select Case n
                        Case "DESAS DESUS", "SENSASIHOT"
                            n = "INFOTAINMENT"
                        Case "CINTA FITRI SEASON 2", "SAMUEL"
                            n = "SERIES1"
                        Case "TERLANJUR INDAH"
                            n = "SERIES3"
                        Case "CINTA CINDERELLA"
                            n = "SERIES2"
                        Case "DUNIA TANPA TUHAN"
                            n = "SERIES4"
                        ' Jika tidak ada kecocokan, biarkan n tetap sesuai nilai A4
                    End Select
                Else
                    ' Jika D4 bukan "MDTV", ubah nama sheet menjadi "KOMPETITOR"
                    n = "KOMPETITOR"
                End If
                
                ' Cek apakah nama sheet valid
                If n <> "" Then
                    Done = False
                    On Error Resume Next
                    Do Until Done
                        ws.Name = n
                        n = n & " "
                        Done = (Err.Number = 0)
                    Loop
                    On Error GoTo 0
                    ws.Visible = xlSheetVisible
                Else
                    ' Sembunyikan sheet jika tidak memenuhi kriteria
                    ws.Visible = xlSheetHidden
                End If
            End If
        End If
        i = i + 1
    Next ws
    
    ' Kembali ke sheet pertama
    wb.Sheets(1).Select
    wb.Sheets(1).Range("A1").Select
End Sub