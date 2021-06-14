
Sub call_for_all_sheets()
    Dim curr As Worksheet
    For Each curr In ThisWorkbook.Worksheets
        With curr
        .Activate
            Call VBAChallenge
        End With
    Next
End Sub

Sub VBAChallenge()
Application.ScreenUpdating = False

Dim i As Long
Dim tick As String
Dim cell As Range
Dim w As Range
Dim cell2 As Long
Dim temprow As Long
Dim temprow1 As Long

Cells(1, 9).Value = "TICKER"
Cells(1, 10).Value = "YEARLY CHANGE"
Cells(1, 11).Value = "PERCENT CHANGE"
Cells(1, 12).Value = "TOTAL STOCK VOLUME"
'Below headers part of extra challenge
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

' Get all unique tickers in new column
For i = 2 To Rows.Count
    Cells(i, 9).Value = Cells(i, 1).Value

Next i
Range("I:I").RemoveDuplicates Columns:=1, Header:=xlYes

Set cell = Range("A:A")
cell2 = Range("I1").End(xlDown).Row
'Find metrics
For i = 2 To cell2
    tick = Cells(i, 9).Value
    Set w = cell.Find(what:=tick, after:=cell(1), lookat:=xlWhole, searchdirection:=xlPrevious)
    temprow = Mid(w.Address(0, 0), 2)
    ' temprow is row number of last instance of given ticker, i.e. last date
    Set w = cell.Find(what:=tick, after:=cell(1), lookat:=xlWhole)
    ' temprow1 gives row of first instance of ticker
    temprow1 = Mid(w.Address(0, 0), 2)
    Cells(i, 10).Value = Cells(temprow, 6).Value - Cells(temprow1, 3).Value
    'Percent Change
    On Error Resume Next
        If IsError((Cells(i, 10).Value / Cells(temprow1, 3).Value)) Then
            Cells(i, 11).Value = 0
        Else
            Cells(i, 11).Value = (Cells(i, 10).Value / Cells(temprow1, 3).Value)
        End If
    On Error GoTo 0
    ' Total stock volume
    Cells(i, 12).Value = Application.WorksheetFunction.SumIf(Range("A:A"), tick, Range("G:G"))
Next i
' Conditional formatting on yearly change column and percent change
Range(Cells(2, 10), Cells(cell2, 10)).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
Range(Cells(2, 10), Cells(cell2, 10)).FormatConditions(1).Interior.Color = RGB(255, 0, 0)
Range(Cells(2, 10), Cells(cell2, 10)).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
Range(Cells(2, 10), Cells(cell2, 10)).FormatConditions(2).Interior.Color = RGB(0, 255, 0)
'Range("J1").Interior.Color = xlNone
Range("K:K").NumberFormat = "0.00%"

'Find greatest values in various categories
Cells(2, 17).Value = Application.WorksheetFunction.Max(Range("K:K"))
Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("K:K"))
Range(Cells(2, 17), Cells(3, 17)).NumberFormat = "0.00%"
Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("L:L"))
' Implementing a faster version of VLookUP
'Source https://www.mrexcel.com/board/threads/vlookup-vba-alternative.1043185/
inarr = Range(Cells(1, 9), Cells(cell2, 12))
'greatest % increase
want = Cells(2, 17)

For i = 1 To cell2
    If inarr(i, 3) = want Then
        'Debug.Print "in if"
        Cells(2, 16) = inarr(i, 1)
        Exit For
    End If
Next i
' greatest % decrease

want = Cells(3, 17)
For i = 1 To cell2
    If inarr(i, 3) = want Then
        Cells(3, 16) = inarr(i, 1)
        Exit For
    End If
Next i
' greatest stock vol
want = Cells(4, 17)
For i = 1 To cell2
    If inarr(i, 4) = want Then
        Cells(4, 16) = inarr(i, 1)
        Exit For
    End If
Next i


Application.ScreenUpdating = True

ActiveSheet.Cells.EntireColumn.AutoFit
End Sub

