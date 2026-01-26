Sub Macro1()

Application.ScreenUpdating = False


'
' Macro1 Macro
'

'
    Sheets("dp").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("data").Select
    Range("C6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("dp").Select
    Range("C5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("data").Select
    Range("E6").Select
    ActiveWindow.SmallScroll Down:=-9
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("prog").Select
    
    
    Sheets("data").Select
    
'
awal = 5
Do Until Cells(awal, 2) = ""
    Sheets("prog").Select
    stime = Trim(Left(Cells(awal, 2), 5))
    prog = Cells(awal, 3)

    Sheets("data").Select
    brs = 6
    Do Until Left(Cells(brs, 3), 5) = stime
        brs = brs + 1
    Loop
    
    Cells(brs, 4) = prog

awal = awal + 1
Loop

For j = 6 To 911

If Cells(j, 3) <> "" And Cells(j, 4) = "" Then
    nama_data = Cells(j - 1, 4)
    Cells(j, 4) = nama_data
Else
    End If
Next j
 
    
End Sub
