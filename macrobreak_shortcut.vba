Sub Break()
    ' Break Macro
    ' Keyboard Shortcut: Ctrl + Shift + Q

    On Error Resume Next
    Columns("C:C").Delete Shift:=xlToLeft

    Cells.Replace What:="Day Part", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False

    Dim a As Range
    For Each a In Range("B6", Range("B" & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeConstants).Areas
        a.Offset(1).ClearContents
        a.Cells(a.Rows.Count + IIf(a.Rows.Count = 1, 1, 0)) = "_1____#1"
    Next a

    Range("A1").Select
End Sub
Sub Auto_Open()
    ' Daftarkan shortcut saat Excel dibuka
    Application.OnKey "^+q", "Break"   ' Ctrl + Shift + Q
End Sub

Sub Auto_Close()
    ' Hapus shortcut saat Excel ditutup
    Application.OnKey "^+q"
End Sub

