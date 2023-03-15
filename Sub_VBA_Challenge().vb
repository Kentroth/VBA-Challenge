Sub VBA_Challenge()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

ws.Range("A:A").Copy
ws.Range("I:I").PasteSpecial
    
ws.Range("I:I").RemoveDuplicates Columns:=1, Header:=xlYes

ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("l1").Value = "Total Stock Volume"

ws.Range("I1:L1").Columns.AutoFit

Dim LR As Integer
 LR = ws.Cells(Rows.Count, 9).End(xlUp).Row

For r = 2 To LR
        
    Dim StartRow As Long, EndRow As Long
    With ws
        StartRow = Range("A:A").Find(what:=Cells(r, 9), after:=Range("A1")).Row
        EndRow = Range("A:A").Find(what:=Cells(r, 9), after:=Range("A1"), searchdirection:=xlPrevious).Row
    End With

    ws.Cells(r, 10) = ws.Cells(EndRow, 6) - ws.Cells(StartRow, 3)
    ws.Range("J:J").NumberFormat = "0.00"
    
    If ws.Cells(r, 10) < 0 Then
        ws.Cells(r, 10).Interior.ColorIndex = 3
        Else
        ws.Cells(r, 10).Interior.ColorIndex = 4
    End If
        
    ws.Cells(r, 11) = (ws.Cells(EndRow, 6) - ws.Cells(StartRow, 3)) / (ws.Cells(StartRow, 3) + 0.00000000000001)
    ws.Range("K:K").NumberFormat = "0.00%"
    
    Dim rngcriteria As Range
    Dim rngsum As Range
    Set rngcriteria = ws.Range("A:A")
    Set rngsum = ws.Range("G:G")
    
    ws.Cells(r, 12) = WorksheetFunction.SumIf(rngcriteria, ws.Cells(r, 9), rngsum)

Next r

ws.Range("o2") = "Greatest % Increase"
ws.Range("o3") = "Greatest % Decrease"
ws.Range("o4") = "Greatest Total Volume"
ws.Range("p1") = "Ticker"
ws.Range("q1") = "Value"

ws.Range("P2") = WorksheetFunction.XLookup(WorksheetFunction.Max(ws.Range("j:j")), ws.Range("J:J"), ws.Range("I:I"))
ws.Range("P3") = WorksheetFunction.XLookup(WorksheetFunction.Min(ws.Range("j:j")), ws.Range("J:J"), ws.Range("I:I"))
ws.Range("P4") = WorksheetFunction.XLookup(WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), ws.Range("I:I"))

ws.Range("Q2") = WorksheetFunction.Max(ws.Range("j:j"))
ws.Range("Q3") = WorksheetFunction.Min(ws.Range("j:j"))
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L:L"))

ws.Range("o1:q4").Columns.AutoFit

Next ws


End Sub
