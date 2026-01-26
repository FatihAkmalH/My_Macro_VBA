
Sub macro_mbym_newed()
    Application.ScreenUpdating = False

    ' Path yang sudah benar (network drive Z)
    Dim folder1 As String, folder2 As String
    folder1 = "O:\DEVELOPMENT\Raw data\MbyM\"
    folder2 = Cells(9, 6)

    ' Cek apakah workbook utama sudah terbuka
    If Not WorkbookIsOpen("Macro MbyM MDTV - Jam 10 (TotalTV) - ARIANNA 29055.xlsm") Then
        MsgBox "Workbook utama belum dibuka!", vbExclamation
        Exit Sub
    End If

    Windows("Macro MbyM MDTV - Jam 10 (TotalTV) - ARIANNA 29055.xlsm").Activate
    Sheets("Makro").Select

    Dim datacase As String, dailydate As String
    datacase = Cells(6, 6)
    dailydate = Cells(10, 6)

    Dim file1 As String, file2 As String, file3 As String, file4 As String
    file1 = "Prg " & dailydate & ".xls"
    file2 = "Break " & dailydate & ".xls"
    file3 = "MbyM " & dailydate & ".xls"
    file4 = "SerCom " & dailydate & ".xls"

    Dim mulai As Long
    mulai = 12

    Do While Cells(mulai, 6) <> "" And Cells(mulai, 7) <> ""
        Dim namafile As String, prgfilter As String, weeknow As String, chn As String
        namafile = Cells(mulai, 6)
        prgfilter = Cells(mulai, 7)
        weeknow = Left(namafile, 4)
        chn = Cells(mulai, 8)

        ' Call macro dengan folder yang sudah benar
        Call Edit_Source(weeknow, folder1)
        Call Create_MbyM(weeknow, folder2, chn, prgfilter, namafile)
        Call CopyAngkaPivot(weeknow, folder2, chn, prgfilter, namafile)
        Call LeadInOut(weeknow, folder2, chn, prgfilter, namafile, dailydate, folder1)

        mulai = mulai + 1

        ' Tutup file jika terbuka
        If WorkbookIsOpen(file1) Then Workbooks(file1).Close SaveChanges:=False
        If WorkbookIsOpen(file3) Then Workbooks(file3).Close SaveChanges:=False
        If WorkbookIsOpen(file4) Then Workbooks(file4).Close SaveChanges:=False
    Loop

    Application.ScreenUpdating = True

    Windows("Macro MbyM MDTV - Jam 10 (TotalTV) - ARIANNA 29055.xlsm").Activate
    Sheets("Makro").Select
End Sub

Function WorkbookIsOpen(wbName As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(wbName)
    WorkbookIsOpen = Not wb Is Nothing
    On Error GoTo 0
End Function

Sub Edit_Source(weeknow, folder1)
Application.ScreenUpdating = False
Windows("Macro MbyM MDTV - Jam 10 (TotalTV) - ARIANNA 29055.xlsm").Activate
Sheets("Makro").Select
        datacase = Cells(6, 6)
        dailydate = Cells(10, 6)
        Select Case UCase(datacase)
'            Case "WEEKLY"
'                file1 = "Prg " & weeknow & ".xls"
'                file2 = "Break " & weeknow & ".xls"
'                file3 = "MbyM " & weeknow & ".xls"
'                folder1 = "R:\Export-Ariana\Weekly\MByM\"
            Case "DAILY"
                file1 = "Prg " & dailydate & ".xls"
'                file2 = "Break " & dailydate & ".xls"
                file3 = "MbyM " & dailydate & ".xls"
                file4 = "SerCom " & dailydate & ".xls"
                folder1 = "O:\DEVELOPMENT\Raw data\MbyM\"
        End Select
    Application.DisplayAlerts = False
    Workbooks.Open FileName:=folder1 & file1
'    Workbooks.Open FileName:=folder1 & file2
    Workbooks.Open FileName:=folder1 & file3
    Workbooks.Open FileName:=folder1 & file4
    Application.DisplayAlerts = True
    
    Windows(file1).Activate
    If Cells(1, 1) <> "EDITED" And Cells(1, 2) = "" Then
        Range(Cells(1, 1), Cells(1, 20)).Select
        Selection.EntireColumn.MergeCells = False
        Selection.EntireColumn.Interior.ColorIndex = xlNone
        Cells(1, 1).Select
             
             For x1 = 1 To 50
                cekx1 = Trim(Cells(x1, 1))
                If Left(cekx1, 15) = "Selected target" Then
                    brsuniverse = x1 + 1
                    Exit For
                End If
            Next x1
        If brsuniverse = Empty Then brsuniverse = 11
        
        data = Cells(brsuniverse, 1)
        data = extract(data, 2, ":")
        data = extract(data, 1, " ")
        Cells(1, 1) = "EDITED"
        Cells(1, 2) = data
        universe = data
        For i = 1 To 50
            If Cells(i, 1) = "Counter" Then
                brsmulai = i
                Exit For
            End If
        Next i
        If brsmulai = "" Then brsmulai = 14
        For i = brsmulai + 1 To 65000
            If Cells(i, 1) <> "" Then
                If Cells(i, 2) = "" Then Cells(i, 2) = Cells(i - 1, 2)
                If Cells(i, 3) <> "" Then
                    Cells(i, 3) = Left(Cells(i, 3), 3)
                Else
                    Cells(i, 3) = Cells(i - 1, 3)
                End If
            Else
                Exit For
            End If
        Next i
        Workbooks(file1).Save
    End If
    
    ' lakukan untuk file 2
'        Windows(file2).Activate
'        For sht = 1 To 15
'            If SheetExists(sht) Then
'                Sheets(sht).Select
'                If Cells(1, 1) <> "EDITED" Then
'                    Range(Cells(1, 1), Cells(1, 20)).Select
'                    Selection.EntireColumn.Select
'                    Selection.MergeCells = False
'                    Selection.Interior.ColorIndex = xlNone
'                    Cells(1, 1).Select
'                    brsmulai = ""
'                    For i = 1 To 50
'                        If Trim(Cells(i, 3)) = "Channel" Then
'                            brsmulai = i + 2        ' menentukan baris awal
'                            Exit For
'                        End If
'                    Next i
'                    If brsmulai = "" Then brsmulai = 15     ' default 15, cek bila ada perubahan
'
'                    namasts = Cells(brsmulai - 1, 3)
'                    ActiveSheet.Name = namasts
'                    For i = brsmulai + 1 To 65000
'                        If Cells(i, 4) <> "" Then
'                            If Cells(i, 1) <> "" Then
'                                Cells(i, 1) = Left(Cells(i, 1), 3)
'                            Else
'                                Cells(i, 1) = Cells(i - 1, 1)
'                                Cells(i, 2) = Cells(i - 1, 2)
'                            End If
'                        Else
'                            Exit For
'                        End If
'                    Next i
'                    For i = brsmulai + 1 To 65000
'                        If Cells(i, 4) <> "" Then
'                            wkawal = Left(Cells(i, 3), 5)
'                            datadur = Cells(i, 4)
'                            If datadur > 25 Then
'                                durdecimal = datadur / 60
'                                durfixed = Fix(datadur / 60)
'                                If durdecimal - durfixed > 0.4 Then
'                                    durasi = durfixed + 1
'                                Else
'                                    durasi = durfixed
'                                End If
'
'                                Cells(i, 3) = wkawal
'                                If Cells(i, 1) <> "" Then
'                                    Cells(i, 1) = Left(Cells(i, 1), 3)
'                                Else
'                                    Cells(i, 1) = Cells(i - 1, 1)
'                                    Cells(i, 2) = Cells(i - 1, 2)
'                                End If
'
'                                If durasi > 1 Then
'                                    If Cells(i, 1) <> "" Then
'                                        Cells(i, 1) = Left(Cells(i, 1), 3)
'                                    Else
'                                        Cells(i, 1) = Cells(i - 1, 1)
'                                        Cells(i, 2) = Cells(i - 1, 2)
'                                    End If
'
'                                    For tambahbaris = 1 To durasi - 1
'                                        Cells(i + 1, 3).EntireRow.Insert
'                                        Cells(i + 1, 1) = Cells(i, 1)
'                                        Cells(i + 1, 2) = Cells(i, 2)
'                                    Next tambahbaris
'                                    Cells(i, 3) = "'" & wkawal
'                                    valwaktu = Val(Left(wkawal, 2)) * 60 + Val(Right(wkawal, 2))
'                                    For tambahdurasi = i + 1 To i + durasi - 1
'                                        valwaktu = valwaktu + 1
'                                        nilai1 = Fix(valwaktu / 60)
'                                        nilai2 = valwaktu - (nilai1 * 60)
'                                        Cells(tambahdurasi, 3) = "'" & Format(nilai1, "00") & ":" & Format(nilai2, "00")
'                                        Cells(tambahdurasi, 4) = 60
'                                        Cells(tambahdurasi, 5) = Cells(tambahdurasi - 1, 5)
'                                    Next tambahdurasi
'                                    i = i + durasi - 1
'                                End If
'                            Else
'                                Cells(i, 3).EntireRow.Delete
'                                i = i - 1
'                            End If
'                        Else
'                            Exit For
'                        End If
'                    Next i
'                    Cells(1, 1) = "EDITED"
'                End If
'            End If
'        Next sht
'        Workbooks(file2).Save

    ' rapikan file 3
    Windows(file3).Activate
    If Cells(1, 1) <> "EDITED" And Cells(1, 2) = "" Then
        Range(Cells(1, 1), Cells(1, 20)).Select
        Selection.EntireColumn.Select
        Selection.MergeCells = False
        Selection.Interior.ColorIndex = xlNone
        Cells(1, 1).Select
              
            For x1 = 1 To 50
                cekx1 = Trim(Cells(x1, 1))
                If Left(cekx1, 15) = "Selected target" Then
                    brsuniverse = x1 + 1
                    Exit For
                End If
            Next x1
        If brsuniverse = Empty Then brsuniverse = 11
        
        data = Cells(brsuniverse, 1)

        dd = extract(data, 2, ":")
        cekdata = extract(Trim(dd), 1, " ")
        universe = cekdata
        Cells(1, 1) = "EDITED"
        Cells(1, 2) = universe
        mulai = ""
        For i = 1 To 50
            If Trim(UCase(Cells(i, 1))) = "DAY OF WEEK" Then
                mulai = i
                Exit For
            End If
        Next i
        For i = mulai + 1 To 65000
            If Cells(i, 2) <> "" Then
                If Cells(i, 1) <> "" Then
                    data = Cells(i, 1)
                    Cells(i, 1) = Left(data, 3)
                Else
                    Cells(i, 1) = Cells(i - 1, 1)
                End If
            Else
                Exit For
            End If
        Next i
        Workbooks(file3).Save
    End If
    
'    ' rapikan file 4
'    Windows(file4).Activate
'    If Cells(1, 1) <> "EDITED" And Cells(1, 2) = "" Then
'        Range(Cells(1, 1), Cells(1, 20)).Select
'        Selection.EntireColumn.Select
'        Selection.MergeCells = False
'        Selection.Interior.ColorIndex = xlNone
'        Cells(1, 1).Select
'    End If
    
Application.ScreenUpdating = False
Windows("Macro MbyM MDTV - Jam 10 (TotalTV) - ARIANNA 29055.xlsm").Activate
Sheets("Makro").Select
End Sub

Sub Create_MbyM(weeknow, folder2, chn, prgfilter, namafile)
Application.ScreenUpdating = False
    Dim breaking(1 To 1441) As String
    Dim breakchart(1 To 500) As Integer
    Dim segmen(1 To 40, 1 To 3) As Integer
Windows("Macro MbyM MDTV - Jam 10 (TotalTV) - ARIANNA 29055.xlsm").Activate
Sheets("Makro").Select
        datacase = Cells(6, 6)
        dailydate = Cells(10, 6)
        haridantanggal = Cells(8, 6)
        
        Select Case UCase(datacase)
'            Case "WEEKLY"
'                file1 = "Prg " & weeknow & ".xls"
'                file2 = "Break " & weeknow & ".xls"
'                file3 = "MbyM " & weeknow & ".xls"
'                folder1 = "Z:\Export-Ariana\Weekly\MByM\"
            Case "DAILY"
                file1 = "Prg " & dailydate & ".xls"
                file2 = "Break " & dailydate & ".xls"
                file3 = "MbyM " & dailydate & ".xls"
                file4 = "SerCom " & dailydate & ".xls"
                folder1 = "C:\Export\"
        End Select
    
    foldersave = Cells(9, 6)
    filedflt = "MByM-" & UCase(chn) & " (TotalTV).xlsx"
     Application.DisplayAlerts = False
    Workbooks.Open FileName:="O:\DEVELOPMENT\#aws\" & filedflt
     Application.DisplayAlerts = True
    Windows(file1).Activate
    Sheets(1).Select
    mulaiprg = ""
    For i = 1 To 50
        data = Cells(i, 1)
        If Trim(data) = "Counter" Then
            mulaiprg = i + 1
            Exit For
        End If
    Next i
    If mulaiprg = "" Then mulaiprg = 15
    
'    Windows(file2).Activate
'    Sheets(1).Select
'    mulaibreak = ""
'    For i = 1 To 50
'        data = Cells(i, 1)
'        If Trim(data) = "Channel" Then
'            mulaibreak = i + 1
'            Exit For
'        End If
'    Next i
'    If mulaibreak = "" Then mulaibreak = 16
    
    Windows(file3).Activate
    Sheets(1).Select
    mulaimbym = ""
    For i = 1 To 50
        data = Cells(i, 1)
        If UCase(Trim(data)) = "DAY OF WEEK" Then
            mulaimbym = i + 1
            Exit For
        End If
    Next i
    If mulaimbym = "" Then mulaimbym = 16
    
    
    ' cek prg
    jumadaprg = 0
    For hh = 1 To 7
        Windows(file1).Activate
        Sheets(1).Select
        Select Case hh
            Case 1
                hari = "Sun"
            Case 2
                hari = "Mon"
            Case 3
                hari = "Tue"
            Case 4
                hari = "Wed"
            Case 5
                hari = "Thu"
            Case 6
                hari = "Fri"
            Case 7
                hari = "Sat"
        End Select
        
    ' program
        adaprg = False
        For i = mulaiprg To 65000
            If Cells(i, 1) <> "" Then
                datasts = Cells(i, 2)
                datahari = Cells(i, 3)
                dataprg = Cells(i, 4)
                If (Trim(Split(UCase(chn), "(")(0)) = UCase(datasts) And UCase(datahari) = UCase(hari)) And (UCase(dataprg) = UCase(prgfilter)) Then
                    jumprg = 1
                    awalprg = i
                    adaprg = True
                    For j = i + 1 To i + 20
                        If Cells(j, 4) = "" Then
                            jumprg = jumprg + 1
                        Else
                            akhirprg = j - 1
                            Exit For
                        End If
                    Next j
                End If
            Else
                Exit For
            End If
            If adaprg = True Then Exit For
        Next i
        
        If adaprg = True Then
            jumadaprg = jumadaprg + 1   ' hitung jumlah prg yang ada selama seminggu
            
            Tanggal = Cells(awalprg, 5)
            wkawal = Cells(awalprg, 6)
            wkakhir = Cells(akhirprg, 7)
            rating = Cells(awalprg, 9)
            share = Cells(awalprg, 10)
            valwkawal = Val(Left(wkawal, 2)) * 60 + Val(Mid(wkawal, 4, 2))
            valwkakhir = Val(Left(wkakhir, 2)) * 60 + Val(Mid(wkakhir, 4, 2))
            
                If valwkawal < 131 Then  ' jika jam tayang kurang dari 02:10
                    batasawal = Val(Left(wkawal, 2)) * 60 + Val(Mid(wkawal, 4, 2))
                    tanpaawalan = True
                Else
                    batasawal = Val(Left(wkawal, 2)) * 60 + Val(Mid(wkawal, 4, 2)) - 10
                    tanpaawalan = False
                End If
                If valwkakhir > 1549 Then   'jika jam tayang lebih dari 25:49
                    batasakhir = Val(Left(wkakhir, 2)) * 60 + Val(Mid(wkakhir, 4, 2))
                    tanpaakhiran = True
                Else
                    batasakhir = Val(Left(wkakhir, 2)) * 60 + Val(Mid(wkakhir, 4, 2)) + 9
                    tanpaakhiran = False
                End If
                
            Windows(file3).Activate
            Sheets(1).Select
            universe = Cells(1, 2)
            awalmbym = ""
            akhirmbym = ""
            For i = mulaimbym To 65000
                If Cells(i, 1) <> "" Then
                    datahari = Cells(i, 1)
                    datawaktu = Cells(i, 2)
                    datawaktu = Left(datawaktu, 5)
                    waktu = Val(Left(datawaktu, 2)) * 60 + Val(Right(datawaktu, 2))
                    If waktu = batasawal And UCase(hari) = UCase(datahari) Then
                        awalmbym = i
                    End If
                    If waktu = batasakhir And UCase(hari) = UCase(datahari) Then
                        akhirmbym = i
                    End If
                Else
                    Exit For
                End If
                If awalmbym <> "" And akhirmbym <> "" Then Exit For
            Next i
            
            ' copy data
            Windows(file3).Activate
            Sheets(1).Select
            Range(Cells(awalmbym, 2), Cells(akhirmbym, 2)).Copy
            Windows(filedflt).Activate
            Sheets("data").Select
            Cells(6, 3).Select
            Selection.PasteSpecial Paste:=xlValue
            Application.CutCopyMode = False
            
            Windows(file3).Activate
            Sheets(1).Select
            Range(Cells(awalmbym, 3), Cells(akhirmbym, 3)).Copy
            Windows(filedflt).Activate
            Sheets("data").Select
            Cells(6, 5).Select
            Selection.PasteSpecial Paste:=xlValue
            Application.CutCopyMode = False
            
            ' ganti range cshr prog kompetitor
            Windows(file3).Activate
            Sheets(1).Select
            Range(Cells(awalmbym, 4), Cells(akhirmbym, 16)).Copy
            Windows(filedflt).Activate
            Sheets("data").Select
            Cells(6, 11).Select
            Selection.PasteSpecial Paste:=xlValue
            Application.CutCopyMode = False
            
            Windows(filedflt).Activate
            Sheets("data").Select
            For bb = 6 To 400
                If Cells(bb, 3) <> "" Then
                    Cells(bb, 2) = Tanggal
                Else
                    End If
            Next bb
                   
            
                Windows(filedflt).Activate
                Sheets("data").Select
                For bb = 6 To 500
                    If Cells(bb, 3) = "" Then
                        barisakhir = bb
                        Exit For
                    End If
                Next bb
                
                Range(Cells(barisakhir, 1), Cells(500, 1)).Select
                Selection.EntireRow.Delete
                
                Range(Cells(barisakhir - 10, 1), Cells(barisakhir - 1, 21)).Select
                With Selection.Interior
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.249977111117893
                    .PatternTintAndShade = 0
                End With
                        
    'sercom file 4
    Windows(file4).Activate
    Sheets(1).Select
    
    For cc = 19 To 300
        If Cells(cc, 3) = dataprg Then
            segawal = cc
            Exit For
        Else
        End If
    Next cc
    
    For ee = segawal + 1 To 300
        If Cells(ee, 3) <> dataprg Then
            segakhir = ee - 1
            Exit For
        Else
        End If
    Next ee
    
    Range(Cells(segawal, 17), Cells(segakhir, 17)).Copy

    Windows(filedflt).Activate
    Sheets("SeriesComp").Select
    Range("C32").Select
    Selection.PasteSpecial Paste:=xlValue
    Application.CutCopyMode = False
            
    Windows(file4).Activate
    Sheets(1).Select
    
    Range(Cells(segawal, 31), Cells(segakhir, 31)).Copy
            
    Windows(filedflt).Activate
    Sheets("SeriesComp").Select
    Range("D32").Select
    Selection.PasteSpecial Paste:=xlValue
    Application.CutCopyMode = False
            
    Windows(file4).Activate
    Sheets(1).Select
    
    Range(Cells(segawal, 45), Cells(segakhir, 45)).Copy
            
    Windows(filedflt).Activate
    Sheets("SeriesComp").Select
    Range("E32").Select
    Selection.PasteSpecial Paste:=xlValue
    Application.CutCopyMode = False
            
    Windows(file4).Activate
    Sheets(1).Select
    
    Range(Cells(segawal, 59), Cells(segakhir, 59)).Copy
            
    Windows(filedflt).Activate
    Sheets("SeriesComp").Select
    Range("F32").Select
    Selection.PasteSpecial Paste:=xlValue
    Application.CutCopyMode = False
            
    Windows(file4).Activate
    Sheets(1).Select
    
    Range(Cells(segawal, 73), Cells(segakhir, 73)).Copy
            
    Windows(filedflt).Activate
    Sheets("SeriesComp").Select
    Range("G32").Select
    Selection.PasteSpecial Paste:=xlValue
    Application.CutCopyMode = False
            
    Windows(file4).Activate
    Sheets(1).Select
    
    Range(Cells(segawal, 73), Cells(segakhir, 86)).Copy
            
    Windows(filedflt).Activate
    Sheets("SeriesComp").Select
    Range("AC10").Select
    Selection.PasteSpecial Paste:=xlValue
    Application.CutCopyMode = False
            
            
    Windows(file4).Activate
    Sheets(1).Select
    
    Range(Cells(segawal, 59), Cells(segakhir, 72)).Copy
            
    Windows(filedflt).Activate
    Sheets("SeriesComp").Select
    Range("AC40").Select
    Selection.PasteSpecial Paste:=xlValue
    Application.CutCopyMode = False
            
    'menentukan break
    Windows(filedflt).Activate
    Sheets("SeriesComp").Select
               
    For pp = 32 To 36
        Windows(filedflt).Activate
        Sheets("SeriesComp").Select
        If UCase(Cells(pp + 1, 3)) <> "" Then
            breakawal = Left(Format(Cells(pp, 4), "hh:mm"), 2) * 60 + Mid(Format(Cells(pp, 4), "hh:mm"), 4, 2)
            breakakhir = Left(Format(Cells(pp + 1, 3), "hh:mm"), 2) * 60 + Mid(Format(Cells(pp + 1, 3), "hh:mm"), 4, 2) - 1
            cekakhir = Mid(Cells(pp + 1, 3), 4, 2)
 '               If cekakhir <> "00" Then
 '                   breakakhir = Left(Cells(pp + 1, 3), 2) * 60 + Mid(Cells(pp + 1, 3), 4, 2) - 1
 '               Else
 '                   breakakhir = Left(Cells(pp + 1, 3), 2) - 1 & ":59"
 '               End If
                  
            For gg = breakawal To breakakhir
                Windows(filedflt).Activate
                Sheets("data").Select
                For ii = 6 To 500
                    If Left(Cells(ii, 3), 2) * 60 + Mid(Cells(ii, 3), 4, 2) = gg Then
                        Cells(ii, 4) = "Break"
                        Cells(ii, 6) = 100
                            Cells(ii, 4).Select
                            With Selection.Font
                                .Color = -16776961
                                .TintAndShade = 0
                            End With
                        Exit For
                    End If
                Next ii
            Next gg
            
        Else
            Exit For
        End If
    Next pp
               
               
'    ' cek iklan
'            Windows(file2).Activate
'            If SheetExists(chn) Then
'                Sheets(chn).Select
'                For Z = mulaibreak To 65000
'                    If Cells(Z, 1) = hari Then
'                        awalbreak = Z
'                        Exit For
'                    End If
'                Next Z
'                ' bersihkan data breaking()
'                For valmat = 1 To 1440
'                    breaking(valmat) = ""
'                Next
'
'                Windows(file2).Activate
'                Sheets(chn).Select
'                valmat = 1
'                    valbreak = awalbreak
'                    Do While Cells(valbreak, 1) = hari
'                        ' Cells(valbreak, 1).Select
'                        datawaktu = Left(Cells(valbreak, 3), 5)
'                        If Cells(valbreak, 6) <> "X" Then
'                            breaking(valmat) = datawaktu
'                            valmat = valmat + 1
'                            'Exit Do
'                        End If
'                        valbreak = valbreak + 1
'                    Loop
'
'                Windows(filedflt).Activate
'                Sheets("data").Select
'                If tanpaawalan = False Then
'                    brscekbreak = 17
'                Else
'                    brscekbreak = 7
'                End If
'                For x = brscekbreak To barisakhir - 12
'                    If Cells(x, 3) <> "" Then
'                        datawaktu = Left(Cells(x, 3), 5)
'                        If Cells(x, 5).Borders(xlEdgeLeft).LineStyle <> xlNone Then
'                            For cekbreak = 1 To 1440
'                                If breaking(cekbreak) <> "" Then
'                                    If breaking(cekbreak) = datawaktu Then
'                                        Range(Cells(x, 3), Cells(x, 3)).Interior.ColorIndex = 43    ' MENENTUKAN WARNA BREAK
'                                        Cells(x, 4) = "Break"
'                                        Cells(x, 6) = 30
'                                            Cells(x, 4).Select
'                                            With Selection.Font
'                                                .Color = -16776961
'                                                .TintAndShade = 0
'                                            End With
'                                        Exit For
'                                    End If
'                                Else
'                                    Exit For
'                                End If
'                            Next cekbreak
'                        End If
'                    Else
'                        Exit For
'                    End If
'                Next x
'            End If
                                                        
            
'            ' breaking competitor
'            Windows(filedflt).Activate
'            Sheets("data").Select
'
'            For kolsts = 9 To 17
'                ' bersihkan data breaking()
'                For valmat = 1 To 1440
'                    breaking(valmat) = ""
'                Next
'                breaksts = Cells(5, kolsts)
'
'                Windows(file2).Activate
'                If SheetExists(breaksts) Then
'                    Sheets(breaksts).Select
'                    For Z = mulaibreak To 65000
'                        If Cells(Z, 1) = hari Then
'                            awalbreak = Z
'                            Exit For
'                        End If
'                    Next Z
'
'                    x = 1
'                    valbreak = awalbreak
'                    Do While Cells(valbreak, 1) = hari
'                        waktu = Left(Cells(valbreak, 3), 5)
'                            If Cells(valbreak, 6) <> "X" Then
'                                breaking(x) = waktu
'                                x = x + 1
'                                'Exit Do
'                            End If
'                        valbreak = valbreak + 1
'                    Loop
'
'                    Windows(filedflt).Activate
'                    Sheets("data").Select
'                    For x = 17 To barisakhir - 12
'                        If Cells(x, kolsts) <> "" Then
'                            datawaktu = Left(Cells(x, 3), 5)
'                                For y = 1 To 1440
'                                    If breaking(y) <> "" Then
'                                        If breaking(y) = datawaktu Then
'                                            Range(Cells(x, kolsts), Cells(x, kolsts)).Interior.ColorIndex = 43    ' MENENTUKAN WARNA BREAK
'                                        End If
'                                    Else
'                                        Exit For
'                                    End If
'                                Next y
'                        End If
'                    Next x
'                End If
'            Next kolsts
                    
       '--------------
            
            
            
            
'            Windows(filedflt).Activate
'            Sheets("data").Select
'
'            For kk = 16 To barisakhir - 11
'                If Cells(kk, 5) > share And Cells(kk, 4) <> "Break" Then
'                    Range(Cells(kk, 4), Cells(kk, 5)).Select
'                    With Selection.Interior
'                        .Pattern = xlSolid
'                        .PatternColorIndex = xlAutomatic
'                        .ThemeColor = xlThemeColorDark2
'                        .TintAndShade = -0.249977111117893
'                        .PatternTintAndShade = 0
'                    End With
'                End If
'            Next kk
                   
                       
            Windows(filedflt).Activate
            Sheets("data").Select
            
            Cells(1, 1) = dataprg
            Cells(2, 1) = haridantanggal & " | " & Left(wkawal, 5) & " - " & Left(wkakhir, 5)
            Cells(3, 1) = "Based On: Data National Urban, TA: F15+ UM; TVR/CShare: " & rating & "/" & share
                    
                    
            Windows(filedflt).Activate
            Sheets("SeriesComp").Select
            
            Range("B4") = dataprg
            Range("B6") = haridantanggal
            
            'hari2 = Format(Tanggal, "dddd")
           ' Windows(filedflt).Activate
            'Sheets(1).Name = hari2
                    

        End If
        Application.DisplayAlerts = True
    Next hh
    ' save as
    If jumadaprg <> 0 Then
        Workbooks(filedflt).SaveAs FileName:=foldersave & namafile & ".xlsx"
        Workbooks(namafile & ".xlsx").Close
    End If
    
Application.ScreenUpdating = False
Windows("Macro MbyM MDTV - Jam 10 (TotalTV) - ARIANNA 29055.xlsm").Activate
Sheets("Makro").Select
End Sub

Function extract(txt, n, separator)
jum = 0
temp = ""
txt = txt & separator
For i = 1 To Len(txt)
    If Mid(txt, i, 1) = separator Then
        jum = jum + 1
        If jum = n Then
            extract = temp
            Exit Function
        Else
            temp = ""
        End If
    Else
        temp = temp & Mid(txt, i, 1)
    End If
Next i
extract = ""
End Function
Function FileExists(FileName$) As Integer
On Error GoTo FileError
    x = FileLen(FileName$)
    FileExists = True
Exit Function
    
FileError:
    FileExists = False
    Exit Function
End Function
Private Function SheetExists(sname) As Boolean
    Dim x As Object
    On Error Resume Next
    Set x = ActiveWorkbook.Sheets(sname)
    If Err = 0 Then SheetExists = True Else SheetExists = False
End Function
' ambil baris terakhir program untuk makro pivot
Sub CopyAngkaPivot(weeknow, folder2, chn, prgfilter, namafile)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim lastRow As Long
    Dim targetRow As Long
    Dim eleventhValue As Variant
    Dim filedflt As String

    filedflt = namafile & ".xlsx"

    Application.DisplayAlerts = False

    ' Buka file sesuai folder2 yang sudah fix
    Set wb = Workbooks.Open(FileName:=folder2 & filedflt)
    Set ws = wb.Sheets("data")

    ' Cari baris terakhir kolom A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    targetRow = lastRow - 10

    If targetRow >= 1 Then
        eleventhValue = ws.Cells(targetRow, "A").Value
        ws.Range("L2").Value = eleventhValue
        ws.Range("L2").Font.Color = RGB(255, 255, 255)
    End If

    wb.Save
    wb.Close SaveChanges:=False

    Application.DisplayAlerts = True
End Sub

Sub LeadInOut(weeknow, folder2, chn, prgfilter, namafile, dailydate, folder1)
    Dim wbSource As Workbook, wsSource As Worksheet
    Dim wbTarget As Workbook, wsTarget As Worksheet, wsOutput As Worksheet
    Dim pathSource As String, pathTarget As String
    Dim i As Long, lastRow As Long, deskripsiMatchRow As Long
    Dim namaDicari As String, resultSebelum As String, resultSesudah As String
    Dim rowSebelum As Long, rowSesudah As Long
    Dim colTVR As Long, colCSHR As Long
    Dim headerRow As Range
    Dim sh As Worksheet
    Dim sheetFound As Boolean

    ' Path file
    pathSource = folder1 & "Prg " & dailydate & ".xls"
    pathTarget = folder2 & namafile & ".xlsx"

    ' Cek file exist
    If Dir(pathTarget) = "" Then
        MsgBox "File target tidak ditemukan: " & pathTarget, vbExclamation
        Exit Sub
    End If
    If Dir(pathSource) = "" Then
        MsgBox "File source tidak ditemukan: " & pathSource, vbExclamation
        Exit Sub
    End If

    ' Buka target & source
    Set wbTarget = Workbooks.Open(pathTarget)
    Set wsTarget = wbTarget.Sheets("data")
    namaDicari = wsTarget.Range("A1").Value
    
    Set wbSource = Workbooks.Open(pathSource)
    sheetFound = False
    For Each sh In wbSource.Sheets
        If InStr(sh.Name, "Top Program by Name_") > 0 Then
            Set wsSource = sh
            sheetFound = True
            Exit For
        End If
    Next sh
    
    If Not sheetFound Then
        MsgBox "Sheet 'Top Program by Name_' tidak ditemukan.", vbExclamation
        wbSource.Close False
        wbTarget.Close False
        Exit Sub
    End If
    
    ' Cari kolom TVR & CSHR
    Set headerRow = wsSource.Range("A17:Z17")
    colTVR = 0: colCSHR = 0
    For i = 1 To headerRow.Columns.Count
        If Trim(UCase(headerRow.Cells(1, i).Value)) = "TVR" Then colTVR = i
        If InStr(UCase(headerRow.Cells(1, i).Value), "CSHR") > 0 Then colCSHR = i
    Next i
    If colTVR = 0 Or colCSHR = 0 Then
        MsgBox "Kolom TVR atau CSHR tidak ditemukan.", vbExclamation
        wbSource.Close False
        wbTarget.Close False
        Exit Sub
    End If
    
    ' Cari baris program
    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
    deskripsiMatchRow = 0
    For i = 18 To lastRow
        If wsSource.Cells(i, 4).Value = namaDicari Then
            deskripsiMatchRow = i
            Exit For
        End If
    Next i
    
    If deskripsiMatchRow = 0 Then
        MsgBox "Program tidak ditemukan: " & namaDicari, vbExclamation
        wbSource.Close False
        wbTarget.Close False
        Exit Sub
    End If
    
    ' Tentukan row sebelumnya & sesudahnya
    rowSebelum = deskripsiMatchRow - 1
    rowSesudah = deskripsiMatchRow + 1

    Dim namaSebelum As String, namaSesudah As String

    If rowSebelum >= 18 Then
        namaSebelum = wsSource.Cells(rowSebelum, 4).Value
        If InStr(UCase(namaSebelum), "SINEMA") > 0 Or InStr(UCase(namaSebelum), "661C") = 1 Then
            namaSebelum = "MD CERITA NYATA"
        ElseIf InStr(UCase(namaSebelum), "SINEMA PAGI") > 0 Or InStr(UCase(namaSebelum), "661E") = 1 Then
            namaSebelum = "MD CERITA NYATA PAGI"
        ElseIf InStr(UCase(namaSebelum), "SINEMA PAGI") > 0 Or InStr(UCase(namaSebelum), "661D") = 1 Then
            namaSebelum = "MD CERITA NYATA PAGI"
        End If
        resultSebelum = namaSebelum & vbLf & _
            "(" & Format(wsSource.Cells(rowSebelum, colTVR).Value, "0.00") & "/" & _
            Format(wsSource.Cells(rowSebelum, colCSHR).Value, "0.0") & ")"
    Else
        resultSebelum = "(Tidak ada data sebelumnya)"
    End If
    
    If rowSesudah <= lastRow Then
        namaSesudah = wsSource.Cells(rowSesudah, 4).Value
        If InStr(UCase(namaSesudah), "SINEMA") > 0 Or InStr(UCase(namaSesudah), "661C") = 1 Then
            namaSesudah = "MD CERITA NYATA"
        ElseIf InStr(UCase(namaSesudah), "SINEMA PAGI") > 0 Or InStr(UCase(namaSesudah), "661E") = 1 Then
            namaSesudah = "MD CERITA NYATA PAGI"
        ElseIf InStr(UCase(namaSesudah), "SINEMA PAGI") > 0 Or InStr(UCase(namaSesudah), "661D") = 1 Then
            namaSesudah = "MD CERITA NYATA PAGI"
        End If
        resultSesudah = namaSesudah & vbLf & _
            "(" & Format(wsSource.Cells(rowSesudah, colTVR).Value, "0.00") & "/" & _
            Format(wsSource.Cells(rowSesudah, colCSHR).Value, "0.0") & ")"
    Else
        resultSesudah = "(Tidak ada data sesudahnya)"
    End If
    
    ' Paste ke target
    Set wsOutput = wbTarget.Sheets("AVG SEG")
    wsOutput.Range("F1").Value = resultSebelum
    wsOutput.Range("F2").Value = resultSesudah
    wsOutput.Range("F1:F2").WrapText = True
    
    ' Tutup workbook source
    wbSource.Close False
    
    ' Simpan dan tutup target
    wbTarget.Save
    wbTarget.Close False
    
    ' MsgBox "Data berhasil ditampilkan di F1 & F2.", vbInformation
End Sub