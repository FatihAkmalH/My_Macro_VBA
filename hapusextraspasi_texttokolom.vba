Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    Application.EnableEvents = False

    ' Panggil handlers terpisah
    Call HandleTrimAndDate(Target)
    Call HandleColumnI(Target)

CleanExit:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub


' -------------------------
' Handler TRIM & Format Tanggal (B:XFD)
' -------------------------
Private Sub HandleTrimAndDate(ByVal Target As Range)
    Dim rng As Range
    Dim c As Range
    Dim txt As String

    Set rng = Intersect(Target, Me.Range("B:XFD"))
    If rng Is Nothing Then Exit Sub

    For Each c In rng.Cells
        If IsError(c.Value) Then GoTo NextTrimCell

        txt = CStr(c.Value)

        If Len(txt) > 0 Then
            If IsDate(c.Value) Then
                c.NumberFormat = "dd/mm/yyyy"
            Else
                txt = Replace(txt, Chr(160), " ")
                txt = Trim(txt)
                On Error Resume Next
                txt = Application.WorksheetFunction.Trim(txt)
                On Error GoTo 0

                Do While InStr(txt, "  ") > 0
                    txt = Replace(txt, "  ", " ")
                Loop

                If CStr(c.Value) <> txt Then c.Value = txt
            End If
        End If
NextTrimCell:
    Next c
End Sub


' -------------------------
' Handler Kolom I (TRIM + ganti "dan" + TextToColumns)
' -------------------------
Private Sub HandleColumnI(ByVal Target As Range)
    Dim rngI As Range
    Dim c As Range
    Dim txt As String
    Dim rngTrim As Range
    Dim t As Range

    Set rngI = Intersect(Target, Me.Columns("H"))
    If rngI Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    For Each c In rngI.Cells
        If IsError(c.Value) Then GoTo NextI

        txt = CStr(c.Value)
        If Len(txt) > 0 Then
            ' Normalisasi teks awal (kolom H)
            txt = Replace(txt, Chr(160), " ")
            txt = Trim(txt)
            Do While InStr(txt, "  ") > 0
                txt = Replace(txt, "  ", " ")
            Loop

            txt = Replace(txt, " dan ", ", ")
            txt = Replace(txt, " dan", ", ")
            txt = Replace(txt, "dan ", ", ")

            If CStr(c.Value) <> txt Then c.Value = txt

            Application.DisplayAlerts = False

            ' Text to Columns ke kolom N
            c.TextToColumns _
                Destination:=Me.Cells(c.Row, "N"), _
                DataType:=xlDelimited, _
                TextQualifier:=xlTextQualifierDoubleQuote, _
                ConsecutiveDelimiter:=True, _
                Comma:=True, _
                Other:=True, OtherChar:="&"

            Application.DisplayAlerts = True

            ' === TRIM HASIL TextToColumns (N - W) ===
            Set rngTrim = Me.Range(Me.Cells(c.Row, "N"), Me.Cells(c.Row, "W"))

            For Each t In rngTrim.Cells
                If Not IsError(t.Value) Then
                    If VarType(t.Value) = vbString Then
                        t.Value = Replace(t.Value, Chr(160), " ")
                        t.Value = Trim(t.Value)
                    End If
                End If
            Next t
        End If
NextI:
    Next c

    ' Application.ScreenUpdating = True
End Sub





