Sub ProfProf2()
    Dim wbProfile As Workbook
    Dim wbTemplate As Workbook
    Dim wsSource As Worksheet
    Dim newSheet As Worksheet
    Dim Weekx As String, Dayx As String, Monthx As String
    Dim folder_savey As String, file_savey As String, fullPath As String
    Dim i As Integer, ws_num As Integer
    Dim baseName As String
    Dim nameCount As Object
    Set nameCount = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Ambil nilai Weekx dari sheet aktif workbook ini
    Weekx = ThisWorkbook.Sheets(1).Cells(10, 5).Value

    ' Buka workbook template dan sumber
    Set wbTemplate = Workbooks.Open("O:\DEVELOPMENT\#aws\Template Profile.xlsx")
    Set wbProfile = Workbooks.Open("C:\Export\Profile.xls")

    ws_num = wbProfile.Worksheets.Count

    ' Loop semua worksheet di Profile.xls sesuai urutan
    For i = 1 To ws_num
        Set wsSource = wbProfile.Worksheets(i)

        ' Copy sheet Template ke posisi paling akhir
        wbTemplate.Sheets("Template").Copy After:=wbTemplate.Sheets(wbTemplate.Sheets.Count)
        Set newSheet = ActiveSheet

        ' Salin data dari Profile.xls ke Template yang baru
        With newSheet
            wsSource.Range("C6:J36").Copy
            .Range("AD6").PasteSpecial Paste:=xlPasteValues

            ' Cek dan ubah nama jika perlu
            Dim valA2 As String
            Dim valA2Key As String
            Dim replacements As Object
            Set replacements = CreateObject("Scripting.Dictionary")

            replacements.Add "SINEMA", "MDTV CERITA NYATA"
            replacements.Add "SINEMA PAGI", "MDTV CERITA NYATA PAGI"
            replacements.Add "SH**TING STAR", "SHOOTING STAR"
            replacements.Add "PROGRESNYA BERAPA PERSEN?", "PROGRESNYA BERAPA PERSEN"
            ' Tambahkan lainnya di sini (key dalam huruf besar)

            valA2 = Trim(wsSource.Range("A2").Value)
            valA2Key = UCase(valA2)

            ' Tambahan kondisi khusus
            If InStr(valA2Key, "661E") > 0 Then
                valA2 = "MDTV CERITA NYATA PAGI"
            ElseIf InStr(valA2Key, "661C") > 0 Then
                valA2 = "MDTV CERITA NYATA"
            ElseIf InStr(valA2Key, "661D") > 0 Then
                valA2 = "MDTV CERITA NYATA PAGI"
            ElseIf replacements.exists(valA2Key) Then
                valA2 = replacements(valA2Key)
            End If

            .Range("Z6").Value = valA2
            .Range("Z4").Value = wsSource.Range("B2").Value
            .Range("AA4").Value = wsSource.Range("C2").Value

            ' Penamaan sheet hasil (dengan cek duplikat dan bersihkan karakter ilegal)
            Dim safeName As String
            baseName = Trim(wsSource.Range("A2").Value)

            If baseName <> "" Then
                safeName = baseName
                safeName = Replace(safeName, "\", "")
                safeName = Replace(safeName, "/", "")
                safeName = Replace(safeName, ":", "")
                safeName = Replace(safeName, "?", "")
                safeName = Replace(safeName, "*", "")
                safeName = Replace(safeName, "[", "")
                safeName = Replace(safeName, "]", "")

                If Len(safeName) > 31 Then safeName = Left(safeName, 31)

                If Not nameCount.exists(safeName) Then
                    nameCount(safeName) = 0
                Else
                    nameCount(safeName) = nameCount(safeName) + 1
                End If

                If nameCount(safeName) = 0 Then
                    .Name = safeName
                Else
                    Dim finalName As String
                    finalName = Left(safeName, 31 - Len(" (x)")) & " (" & nameCount(safeName) & ")"
                    .Name = finalName
                End If
            End If
        End With
    Next i

    ' Hapus sheet Template yang asli
    Dim ws As Worksheet
    For Each ws In wbTemplate.Sheets
        If ws.Name = "Template" Then
            ws.Delete
            Exit For
        End If
    Next ws

    ' Ambil informasi tanggal dari sheet pertama di Profile.xls
    With wbProfile.Sheets(1)
        Dayx = .Cells(2, 6).Value
        Monthx = Left(.Cells(2, 7).Value, 3)
    End With

    ' Susun folder dan nama file
    folder_savey = "O:\DEVELOPMENT\DAILY\" & Weekx & "\3. PROFILE\"
    file_savey = "Profile Prog MDTV " & Dayx & " " & Monthx & ".xlsx"
    fullPath = folder_savey & file_savey

    ' Jika file sudah ada, minta input user untuk nama tambahan
    If Dir(fullPath) <> "" Then
        Dim userSuffix As String
        userSuffix = InputBox("File dengan nama """ & file_savey & """ sudah ada." & vbCrLf & _
                              "Silakan masukkan tambahan nama di belakang:", "Nama File Sudah Ada", "Versi Baru")
        If Trim(userSuffix) = "" Then userSuffix = "Revisi"
        file_savey = "Profile Prog MDTV " & Dayx & " " & Monthx & " (" & userSuffix & ").xlsx"
    End If

    ' Simpan workbook hasil
    wbTemplate.SaveAs FileName:=folder_savey & file_savey

    ' Tutup Profile.xls tanpa menyimpan
    wbProfile.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub ProfJkt()
    Dim wbProfile As Workbook
    Dim wbTemplate As Workbook
    Dim wsSource As Worksheet
    Dim newSheet As Worksheet
    Dim Weekx As String, Dayx As String, Monthx As String
    Dim folder_savey As String, file_savey As String, fullPath As String
    Dim i As Integer, ws_num As Integer
    Dim baseName As String
    Dim nameCount As Object
    Set nameCount = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Ambil nilai Weekx dari sheet aktif workbook ini
    Weekx = ThisWorkbook.Sheets(1).Cells(10, 5).Value

    ' Buka workbook template dan sumber
    Set wbTemplate = Workbooks.Open("O:\DEVELOPMENT\#aws\Template Profile-JKT.xlsx")
    Set wbProfile = Workbooks.Open("C:\Export\Profile JKT.xls")

    ws_num = wbProfile.Worksheets.Count

    ' Loop semua worksheet di Profile.xls sesuai urutan
    For i = 1 To ws_num
        Set wsSource = wbProfile.Worksheets(i)

        ' Copy sheet Template ke posisi paling akhir
        wbTemplate.Sheets("Template").Copy After:=wbTemplate.Sheets(wbTemplate.Sheets.Count)
        Set newSheet = ActiveSheet

        ' Salin data dari Profile.xls ke Template yang baru
        With newSheet
            wsSource.Range("C6:J36").Copy
            .Range("AD6").PasteSpecial Paste:=xlPasteValues

            ' Cek dan ubah nama jika perlu
            Dim valA2 As String
            Dim valA2Key As String
            Dim replacements As Object
            Set replacements = CreateObject("Scripting.Dictionary")

            replacements.Add "SINEMA", "MDTV CERITA NYATA"
            replacements.Add "SINEMA PAGI", "MDTV CERITA NYATA PAGI"
            replacements.Add "SH**TING STAR", "SHOOTING STAR"
            replacements.Add "PROGRESNYA BERAPA PERSEN?", "PROGRESNYA BERAPA PERSEN"
            ' Tambahkan lainnya di sini (key dalam huruf besar)

            valA2 = Trim(wsSource.Range("A2").Value)
            valA2Key = UCase(valA2)

            ' Tambahan kondisi khusus
            If InStr(valA2Key, "661E") > 0 Then
                valA2 = "MDTV CERITA NYATA PAGI"
            ElseIf InStr(valA2Key, "661C") > 0 Then
                valA2 = "MDTV CERITA NYATA"
            ElseIf InStr(valA2Key, "661D") > 0 Then
                valA2 = "MDTV CERITA NYATA"
            ElseIf replacements.exists(valA2Key) Then
                valA2 = replacements(valA2Key)
            End If

            .Range("Z6").Value = valA2
            .Range("Z4").Value = wsSource.Range("B2").Value
            .Range("AA4").Value = wsSource.Range("C2").Value

            ' Penamaan sheet hasil (dengan cek duplikat dan bersihkan karakter ilegal)
            Dim safeName As String
            baseName = Trim(wsSource.Range("A2").Value)

            If baseName <> "" Then
                safeName = baseName
                safeName = Replace(safeName, "\", "")
                safeName = Replace(safeName, "/", "")
                safeName = Replace(safeName, ":", "")
                safeName = Replace(safeName, "?", "")
                safeName = Replace(safeName, "*", "")
                safeName = Replace(safeName, "[", "")
                safeName = Replace(safeName, "]", "")

                If Len(safeName) > 31 Then safeName = Left(safeName, 31)

                If Not nameCount.exists(safeName) Then
                    nameCount(safeName) = 0
                Else
                    nameCount(safeName) = nameCount(safeName) + 1
                End If

                If nameCount(safeName) = 0 Then
                    .Name = safeName
                Else
                    Dim finalName As String
                    finalName = Left(safeName, 31 - Len(" (x)")) & " (" & nameCount(safeName) & ")"
                    .Name = finalName
                End If
            End If
        End With
    Next i

    ' Hapus sheet Template yang asli
    Dim ws As Worksheet
    For Each ws In wbTemplate.Sheets
        If ws.Name = "Template" Then
            ws.Delete
            Exit For
        End If
    Next ws

    ' Ambil informasi tanggal dari sheet pertama di Profile.xls
    With wbProfile.Sheets(1)
        Dayx = .Cells(2, 6).Value
        Monthx = Left(.Cells(2, 7).Value, 3)
    End With

    ' Susun folder dan nama file
    folder_savey = "O:\DEVELOPMENT\DAILY\" & Weekx & "\3. PROFILE\"
    file_savey = "Profile Prog MDTV " & Dayx & " " & Monthx & " (MARKET JKT).xlsx"
    fullPath = folder_savey & file_savey

    ' Jika file sudah ada, minta input user untuk nama tambahan
    If Dir(fullPath) <> "" Then
        Dim userSuffix As String
        userSuffix = InputBox("File dengan nama """ & file_savey & """ sudah ada." & vbCrLf & _
                              "Silakan masukkan tambahan nama di belakang:", "Nama File Sudah Ada", "Versi Baru")
        If Trim(userSuffix) = "" Then userSuffix = "Revisi"
        file_savey = "Profile Prog MDTV " & Dayx & " " & Monthx & " (" & userSuffix & ").xlsx"
    End If

    ' Simpan workbook hasil
    wbTemplate.SaveAs FileName:=folder_savey & file_savey

    ' Tutup Profile.xls tanpa menyimpan
    wbProfile.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub