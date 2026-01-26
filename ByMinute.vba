Sub Auto_Open()
    UserForm2.Show
End Sub
Sub FillColBlanks()

'fill blank cells in column with value above
Dim wks As Worksheet
Dim rng As Range
Dim LastRow As Long
Dim col As Long


Set wks = ActiveSheet
With wks
   col = .Range("P5").Column

   Set rng = .UsedRange  'try to reset the lastcell
   LastRow = .Cells.SpecialCells(xlCellTypeLastCell).Row
   Set rng = Nothing
   On Error Resume Next
   Set rng = .Range(.Cells(7, col), .Cells(LastRow, col)) _
                  .Cells.SpecialCells(xlCellTypeBlanks)
   On Error GoTo 0

   If rng Is Nothing Then
       MsgBox "No blanks found"
       Exit Sub
   Else
       rng.FormulaR1C1 = "=R[-1]C"
   End If

   'replace formulas with values
   With .Cells(5, col).EntireColumn
       .Value = .Value
   End With
   
     
End With

ActiveSheet.Visible = xlVeryHidden

End Sub

Sub Copy()

Worksheets("TVS").Activate

With Worksheets("TVS")
.Range("G3").Copy
.Range("G2").PasteSpecial _
     Paste:=xlPasteValues
End With

Charts("Sunday").Activate

End Sub
Sub ShowDialog()
    UserForm1.Show
End Sub
Sub Copy2()

Worksheets("TVS2").Activate

With Worksheets("TVS2")
.Range("G3").Copy
.Range("G2").PasteSpecial _
     Paste:=xlPasteValues
End With

Charts("Monday").Activate

End Sub

Sub Copy3()

Worksheets("TVS3").Activate

With Worksheets("TVS3")
.Range("G3").Copy
.Range("G2").PasteSpecial _
     Paste:=xlPasteValues
End With

Charts("Tuesday").Activate

End Sub

Sub Copy4()

Worksheets("TVS4").Activate

With Worksheets("TVS4")
.Range("G3").Copy
.Range("G2").PasteSpecial _
     Paste:=xlPasteValues
End With

Charts("Wednesday").Activate

End Sub

Sub Copy5()

Worksheets("TVS5").Activate

With Worksheets("TVS5")
.Range("G3").Copy
.Range("G2").PasteSpecial _
     Paste:=xlPasteValues
End With

Charts("Thursday").Activate

End Sub
Sub Copy6()

Worksheets("TVS6").Activate

With Worksheets("TVS6")
.Range("G3").Copy
.Range("G2").PasteSpecial _
     Paste:=xlPasteValues
End With

Charts("Friday").Activate

End Sub
Sub Copy7()

Worksheets("TVS7").Activate

With Worksheets("TVS7")
.Range("G3").Copy
.Range("G2").PasteSpecial _
     Paste:=xlPasteValues
End With

Charts("Saturday").Activate

End Sub

Sub ImportFileProgram()
    Sheets("Program").Activate
    Range("B3:D60").Select
    Selection.ClearContents
    Range("F3:G60").Select
    Selection.ClearContents
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:D60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program").Activate
        Range("B3").Select
        Sheets("Program").Paste
        
           
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("E4:F60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program").Activate
        Range("F3").Select
        Sheets("Program").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Sunday").Activate
End Sub

Sub ImportFileProgram2()
    Sheets("Program2").Activate
    Range("B3:D60").Select
    Selection.ClearContents
    Range("F3:G60").Select
    Selection.ClearContents
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:D60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program2").Activate
        Range("B3").Select
        Sheets("Program2").Paste
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("E4:F60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program2").Activate
        Range("F3").Select
        Sheets("Program2").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Monday").Activate
End Sub

Sub ImportFileProgram3()
    Sheets("Program3").Activate
    Range("B3:D60").Select
    Selection.ClearContents
    Range("F3:G60").Select
    Selection.ClearContents
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:D60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program3").Activate
        Range("B3").Select
        Sheets("Program3").Paste
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("E4:F60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program3").Activate
        Range("F3").Select
        Sheets("Program3").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Tuesday").Activate
End Sub

Sub ImportFileProgram4()
    Sheets("Program4").Activate
    Range("B3:D60").Select
    Selection.ClearContents
    Range("F3:G60").Select
    Selection.ClearContents
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:D60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program4").Activate
        Range("B3").Select
        Sheets("Program4").Paste
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("E4:F60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program4").Activate
        Range("F3").Select
        Sheets("Program4").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Wednesday").Activate
End Sub

Sub ImportFileProgram5()
    Sheets("Program5").Activate
    Range("B3:D60").Select
    Selection.ClearContents
    Range("F3:G60").Select
    Selection.ClearContents
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:D60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program5").Activate
        Range("B3").Select
        Sheets("Program5").Paste
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("E4:F60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program5").Activate
        Range("F3").Select
        Sheets("Program5").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Thursday").Activate
End Sub

Sub ImportFileProgram6()
    Sheets("Program6").Activate
    Range("B3:D60").Select
    Selection.ClearContents
    Range("F3:G60").Select
    Selection.ClearContents
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:D60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program6").Activate
        Range("B3").Select
        Sheets("Program6").Paste
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("E4:F60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program6").Activate
        Range("F3").Select
        Sheets("Program6").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Friday").Activate
End Sub

Sub ImportFileProgram7()
    Sheets("Program7").Activate
    Range("B3:D60").Select
    Selection.ClearContents
    Range("F3:G60").Select
    Selection.ClearContents
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:D60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program7").Activate
        Range("B3").Select
        Sheets("Program7").Paste
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("E4:F60").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Program7").Activate
        Range("F3").Select
        Sheets("Program7").Paste
        
End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Saturday").Activate
End Sub


Sub ImportFileData()
    
    Sheets("Data").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B5:D1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data").Activate
        Range("C6").Select
        Sheets("Data").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Sunday").Activate
End Sub

Sub ImportFileData2()
    
    Sheets("Data2").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B5:D1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data2").Activate
        Range("C6").Select
        Sheets("Data2").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Monday").Activate
End Sub

Sub ImportFileData3()
    
    Sheets("Data3").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B5:D1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data3").Activate
        Range("C6").Select
        Sheets("Data3").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Tuesday").Activate
End Sub

Sub ImportFileData4()
    
    Sheets("Data4").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B5:D1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data4").Activate
        Range("C6").Select
        Sheets("Data4").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Wednesday").Activate
End Sub

Sub ImportFileData5()
    
    Sheets("Data5").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B5:D1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data5").Activate
        Range("C6").Select
        Sheets("Data5").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Thursday").Activate
End Sub

Sub ImportFileData6()
    
    Sheets("Data6").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B5:D1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data6").Activate
        Range("C6").Select
        Sheets("Data6").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Friday").Activate
End Sub

Sub ImportFileData7()
    
    Sheets("Data7").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B5:D1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data7").Activate
        Range("C6").Select
        Sheets("Data7").Paste

End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Saturday").Activate
End Sub
Sub ImportFileBreak()

    
    Sheets("Data").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B6:D1445").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data").Activate
        Range("P6").Select
        Sheets("Data").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Sunday").Activate
End Sub

Sub ImportFileBreak2()

    
    Sheets("Data2").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B6:D1445").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data2").Activate
        Range("P6").Select
        Sheets("Data2").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Monday").Activate
End Sub

Sub ImportFileBreak3()

    
    Sheets("Data3").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B6:D1445").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data3").Activate
        Range("P6").Select
        Sheets("Data3").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Tuesday").Activate
End Sub

Sub ImportFileBreak4()

    
    Sheets("Data4").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B6:D1445").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data4").Activate
        Range("P6").Select
        Sheets("Data4").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Wednesday").Activate
End Sub

Sub ImportFileBreak5()

    
    Sheets("Data5").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B6:D1445").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data5").Activate
        Range("P6").Select
        Sheets("Data5").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Thursday").Activate
End Sub

Sub ImportFileBreak6()

    
    Sheets("Data6").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B6:D1445").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data6").Activate
        Range("P6").Select
        Sheets("Data6").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Friday").Activate
End Sub

Sub ImportFileBreak7()

    
    Sheets("Data7").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
                
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B6:D1445").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data7").Activate
        Range("P6").Select
        Sheets("Data7").Paste
End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Saturday").Activate
End Sub
Sub ImportFileDataH2H()
    
    Sheets("Data").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:K1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data").Activate
        Range("C5").Select
        Sheets("Data").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Sunday").Activate
End Sub

Sub ImportFileDataH2H2()
    
    Sheets("Data2").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:K1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data2").Activate
        Range("C5").Select
        Sheets("Data2").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Monday").Activate
End Sub

Sub ImportFileDataH2H3()
    
    Sheets("Data3").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:K1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data3").Activate
        Range("C5").Select
        Sheets("Data3").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Tuesday").Activate
End Sub

Sub ImportFileDataH2H4()
    
    Sheets("Data4").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:K1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data4").Activate
        Range("C5").Select
        Sheets("Data4").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Wednesday").Activate
End Sub

Sub ImportFileDataH2H5()
    
    Sheets("Data5").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:K1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data5").Activate
        Range("C5").Select
        Sheets("Data5").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Thursday").Activate
End Sub

Sub ImportFileDataH2H6()
    
    Sheets("Data6").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then
        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:K1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data6").Activate
        Range("C5").Select
        Sheets("Data6").Paste
        
        End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Friday").Activate
End Sub

Sub ImportFileDataH2H7()
    
    Sheets("Data7").Activate
    
    Dim sFile
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    sFile = Application.GetOpenFilename( _
        FileFilter:="All Files (*.*), *.*", FilterIndex:=1, _
        Title:="Import File by Meutia@transtv.co.id")
    If sFile <> False Then

        Workbooks.OpenText Filename:=sFile
        Sheets(1).Activate
        Sheets(1).Range("B4:K1444").Select
        Selection.Copy
        ActiveWorkbook.Close
        Sheets("Data7").Activate
        Range("C5").Select
        Sheets("Data7").Paste

End If
        
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Charts("Saturday").Activate
End Sub


