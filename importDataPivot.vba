Sub ImportData()
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Dim sourceRange As Range
    Dim destinationCell As Range
    Dim sourceFilePath As String
    Dim copyUntilRow As Long
    Dim pasteStartRow As Long
    Dim matchFound As Boolean
    Dim i As Long
    Dim fd As FileDialog
    Dim weekName As String

    ' Tentukan workbook dan sheet tujuan
    Set destinationWorkbook = ThisWorkbook
    Set destinationSheet = destinationWorkbook.Sheets("DATA")

    ' Ambil nilai week dari sel D2 di sheet tujuan
    weekName = destinationSheet.Range("A1").Value

    ' Periksa apakah weekName valid
    If weekName = "" Then Exit Sub

    ' Buat objek FileDialog sebagai dialog pemilihan file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    ' Setel judul dialog
    fd.Title = "Pilih File MBM | Makro by fatih.akmal@mdentertainment.com"

    ' Setel link awal berdasarkan weekName folder
    fd.InitialFileName = "O:\DEVELOPMENT\#HASIL BY MINUTE\PROGRAM WEEK " & weekName & "\#EXCEL BY MINUTE PER DAY\"

    ' Setel filter untuk file Excel
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xls; *.xlsx"

    ' Tampilkan dialog dan pastikan pengguna tidak membatalkan dialog pemilihan file
    If fd.Show = -1 Then
        sourceFilePath = fd.SelectedItems(1)
    Else
        Exit Sub
    End If

    ' Buka workbook sumber (File MBM)
    Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Set sourceSheet = sourceWorkbook.Sheets("data")

    ' Pastikan hanya sheet "data" yang aktif (hindari multiple selected sheets)
    sourceWorkbook.Sheets(sourceSheet.Name).Select

    ' Inisialisasi variabel pencarian
    matchFound = False

    ' Periksa apakah nilai di sel O2 sheet sumber cocok dengan A16:A100 di sheet sumber
    For i = 16 To 180
        If sourceSheet.Range("Z2").Value = sourceSheet.Range("A" & i).Value Then
            matchFound = True
            copyUntilRow = i
            Exit For
        End If
    Next i

    ' Jika kecocokan ditemukan, salin data dari rentang sumber
    If matchFound Then
        Set sourceRange = sourceSheet.Range("B16:V" & copyUntilRow)

        ' Menemukan sel kosong pertama di kolom A sheet tujuan
        With destinationSheet
            pasteStartRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            If .Cells(pasteStartRow, 1).Value <> "" Then
                pasteStartRow = pasteStartRow + 1
            End If
            Set destinationCell = .Cells(pasteStartRow, 1)
        End With

        ' Salin dan paste nilai
        sourceRange.Copy
        destinationCell.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If

    ' Tutup file sumber tanpa menyimpan
    sourceWorkbook.Close SaveChanges:=False
End Sub