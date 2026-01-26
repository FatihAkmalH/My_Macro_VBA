Sub Move_Copy_Format()
    Dim wbTemplate As Workbook
    Dim wsControl As Worksheet
    Dim folderSource As String, folderDest As String, saveFolder As String
    Dim monthFolder As String
    Dim fileName As String, fullPath As String
    Dim wbSource As Workbook
    Dim firstDate As String, lastDate As String
    Dim copiedCount As Long
    Dim wsTemplateSheet As Worksheet
    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim cell As Range
    Dim sheetName As String
    Dim monthName As String, saveName As String
    Dim monthParts() As String
    Dim i As Long
    Dim shp As Shape
    Dim wsNew As Worksheet
    
    On Error Resume Next
    Set wsControl = ThisWorkbook.Sheets(1)
    On Error GoTo 0
    
    If wsControl Is Nothing Then
        MsgBox "Tidak ditemukan Sheet kontrol di file ini.", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' --- Lokasi folder
    folderDest = "Y:\SCHEDULING\MACRO SCHEDULING\" 'lokasi file template.xlsx
    saveFolder = wsControl.Range("F3").Value
    monthFolder = wsControl.Range("B3").Value
    folderSource = "Y:\SCHEDULING\1. Daily schedule\"
    
    ' --- Buka template
    Set wbTemplate = Workbooks.Open(folderDest & "template.xlsx")
    Set wsTemplateSheet = wbTemplate.Sheets("Sheet1")
    
    copiedCount = 0
    
    ' --- Loop file dari D3:D15
    For Each cell In wsControl.Range("D3:D15")
        If Trim(cell.Value) <> "" Then
            fileName = Format(cell.Value, "000000") & ".xls"
            fullPath = folderSource & fileName
            
            If Dir(fullPath) <> "" Then
                Set wbSource = Workbooks.Open(fullPath)
                
                ' Jalankan FormatSheet pada sheet sumber
                Call FormatSheet_Worksheet(wbSource.Sheets(1))
                
                ' Ambil tanggal dari C3 (misal "Sat / 04-Oct-2025")
                sheetName = wbSource.Sheets(1).Range("C3").Value
                If InStr(sheetName, "/") > 0 Then sheetName = Split(sheetName, "/")(1)
                sheetName = Trim(Split(sheetName, "-")(0))
                sheetName = Replace(sheetName, " ", "")
                If IsNumeric(sheetName) Then sheetName = CStr(Val(sheetName))
                
                ' Copy sheet ke template
                wbSource.Sheets(1).Copy After:=wbTemplate.Sheets(wbTemplate.Sheets.Count)
                
                ' Ubah nama sheet di template
                With wbTemplate.Sheets(wbTemplate.Sheets.Count)
                    On Error Resume Next
                    .Name = sheetName
                    If Err.Number <> 0 Then
                        .Name = sheetName & "_" & copiedCount
                        Err.Clear
                    End If
                    On Error GoTo 0
                End With
                
                ' === Copy gambar/logo dari template ===
                Set wsNew = wbTemplate.Sheets(wbTemplate.Sheets.Count)
                For Each shp In wsTemplateSheet.Shapes
                    shp.Copy
                    wsNew.Paste
                    With wsNew.Shapes(wsNew.Shapes.Count)
                        ' Atur posisi kiri & top manual (kiri: 0, turun 6 klik = sekitar 72 poin)
                        .Left = 0
                        .Top = 72
                        .Placement = xlFreeFloating
                    End With
                Next shp
                
                copiedCount = copiedCount + 1
                If copiedCount = 1 Then firstDate = sheetName
                lastDate = sheetName
                
                wbSource.Close SaveChanges:=False
            End If
        End If
    Next cell
    
    ' --- Hide Sheet1 di template
    On Error Resume Next
    wbTemplate.Sheets("Sheet1").Visible = xlSheetHidden
    On Error GoTo 0
    
    ' --- Nama file akhir
    monthName = Trim(wsControl.Range("B6").Value)
    
    If InStr(monthName, "-") > 0 Then
        monthParts = Split(monthName, "-")
        saveName = "Daily Schedule MDTV " & CLng(firstDate) & " " & Trim(monthParts(0)) & _
                   " - " & CLng(lastDate) & " " & Trim(monthParts(1)) & " 2026.xlsx"
    Else
        If copiedCount > 1 Then
            saveName = "Daily Schedule MDTV " & CLng(firstDate) & " - " & CLng(lastDate) & " " & monthName & " 2026.xlsx"
        Else
            saveName = "Daily Schedule MDTV " & CLng(firstDate) & " " & monthName & " 2026.xlsx"
        End If
    End If
    
    ' --- Simpan hasil
    wbTemplate.SaveCopyAs saveFolder & saveName
    Set newWb = Workbooks.Open(saveFolder & saveName)
    wbTemplate.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Proses selesai! " & copiedCount & " sheet diformat & dicopy.", vbInformation
End Sub


'===========================================================
' Fungsi FormatSheet disesuaikan agar bisa dipanggil dari macro lain
'===========================================================
Sub FormatSheet_Worksheet(ws As Worksheet)
    Dim lastRow As Long
    
    On Error Resume Next
    Application.ScreenUpdating = False
    ActiveWindow.DisplayGridlines = False
    
    ws.Activate
    ws.Range("B3:C3,F3:J3").UnMerge
    
    ' Copy sementara
    ws.Range("B3").Copy: ws.Range("A2").PasteSpecial xlPasteValues
    ws.Range("F3").Copy: ws.Range("A1").PasteSpecial xlPasteValues
    
    ' Hapus kolom
    ws.Columns("T:Y").Delete
    ws.Columns("Q:Q").Delete
    ws.Columns("I:K").Delete
    ws.Columns("D:G").Delete
    ws.Columns("B:B").Delete
    
    ' Lebar kolom
    ws.Columns("A").ColumnWidth = 3.14
    ws.Columns("B").ColumnWidth = 10.86
    ws.Columns("C").ColumnWidth = 34.29
    ws.Columns("D").ColumnWidth = 8.14
    ws.Columns("E").ColumnWidth = 12.29
    ws.Columns("F").ColumnWidth = 6
    ws.Columns("G").ColumnWidth = 11.14
    ws.Columns("H").ColumnWidth = 12.86
    ws.Columns("I").ColumnWidth = 54.29
    ws.Columns("J").ColumnWidth = 5.29
    
    ' Header
    With ws.Range("A5:J5")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Last row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 6 Then lastRow = 6
    
    ws.Rows("6:" & lastRow).RowHeight = 33
    
    With ws.Range("A5:J" & lastRow)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("G6:G" & lastRow).HorizontalAlignment = xlCenter
    ws.Range("I6:J" & lastRow).HorizontalAlignment = xlLeft
    
    With ws.Rows("6:" & lastRow).Font
        .Name = "Arial"
        .Size = 10
    End With
    
    ' Kembalikan nilai
    ws.Range("A1").Copy: ws.Range("C3").PasteSpecial xlPasteValues
    ws.Range("A2").Copy: ws.Range("B3").PasteSpecial xlPasteValues
    ws.Range("A1:A2").ClearContents
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub