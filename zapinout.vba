Sub Zap_In_Out()
    '
    ' Zap_In_Out Macro
    '
    
    ' Mengambil nilai dari sel E10 pada worksheet aktif
    Weekx = Cells(10, 5)
    Dayx = Cells(14, 12)

    ' Membuka file yang diperlukan
    Workbooks.Open FileName:="O:\DEVELOPMENT\#aws\Template Zap In Out.xlsm"
    Workbooks.Open FileName:="C:\Export\Zap In Out.xls"
    Windows("Zap In Out.xls").Activate

    ' Menyalin data dari Zap In Out.xls ke Template Zap In Out.xlsm
    Range("B5:CX30").Select
    Selection.Copy
    Windows("Template Zap In Out.xlsm").Activate
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' Panggil sub-makro RenameSheets() untuk mengganti nama sheet sebelum menyimpan file
    Call RenameSheets
    
    ' Menyimpan file Template Zap In Out ke folder yang ditentukan
    folder_savey = "O:\DEVELOPMENT\DAILY\" & Weekx & "\1. ZAP IN OUT\"
    file_savey = "Zap In & Zap Out Week " & Weekx & " (" & Dayx & ")" & " - National Urban" & ".xlsm"
    
    ' Menyimpan workbook ke lokasi dan nama file yang ditentukan
    ActiveWorkbook.SaveAs FileName:=folder_savey & file_savey

    ' Menutup file "Zap In Out.xls" tanpa menyimpan perubahan
    Workbooks("Zap In Out.xls").Close SaveChanges:=False
End Sub
Sub RenameSheets()
    Sheets("Source").Select
    Cells.Replace What:="SH**TING", Replacement:="SHOOTING", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Sheets("Macro").Select

    Dim i As Integer, Done As Boolean
    Dim n As String
    i = 1
    
    ' Loop untuk mengganti nama sheet dan menyembunyikan sheet yang tidak di-rename
    For Each Sheet In Sheets
        ' Jangan sembunyikan sheet "Macro" dan "Source"
        If Sheet.Name <> "Macro" And Sheet.Name <> "Source" Then
            If Cells(i, 1) <> "" Then
                n = Cells(i, 1).Value
                Done = False
                On Error Resume Next
                Do Until Done
                    On Error Resume Next
                    Sheet.Name = n
                    n = n & " "
                    Done = (Err.Number = 0)
                    On Error GoTo 0
                Loop
                ' Sheet yang berhasil di-rename tetap terlihat
                Sheet.Visible = True
                i = i + 1
            Else
                ' Menyembunyikan sheet yang tidak di-rename
                Sheet.Visible = xlSheetHidden
            End If
        Else
            ' Pastikan sheet "Macro" dan "Source" selalu terlihat
            Sheet.Visible = True
        End If
    Next Sheet

    Sheets("Macro").Select
    Range("A1").Select
End Sub