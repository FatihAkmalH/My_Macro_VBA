Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    Dim rngToCheck As Range
    
    Application.EnableEvents = False
    On Error Resume Next
    
    ' Ambil hanya sel di luar kolom A
    Set rngToCheck = Intersect(Target, Me.Range("B:XFD"))
    If rngToCheck Is Nothing Then GoTo ExitHandler
    
    ' Proses sel yang sudah difilter
    For Each cell In rngToCheck
        If Not IsEmpty(cell.Value) Then
            If VarType(cell.Value) = vbString Then
                cell.Value = Application.WorksheetFunction.Trim(cell.Value)
            ElseIf IsDate(cell.Value) Then
                cell.NumberFormat = "dd/mm/yyyy"
            End If
        End If
    Next cell

ExitHandler:
    On Error GoTo 0
    Application.EnableEvents = True
End Sub



