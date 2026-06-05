' Payment Monitoring & Reconciliation Tracker
' Open Payment_Review_Tracker.xlsx in Excel
' Press Alt+F11 > Insert > Module > Paste code as below
' Run macros via Alt+F8


'MACRO 1: Format all sheets 
Sub FormatAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Name = "Arial"
        ws.Cells.Font.Size = 9
        ws.Cells.EntireColumn.AutoFit
    Next ws
    MsgBox "All sheets formatted.", vbInformation, "Done"
End Sub

'MACRO 2: Highlight all HIGH risk claimants 
Sub HighlightHighRisk()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("Claimant Summary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 4 To lastRow
        If ws.Cells(i, 9).Value = "HIGH" Then
            ws.Rows(i).Interior.Color = RGB(255, 199, 199)
            ws.Rows(i).Font.Bold = True
        End If
    Next i
    MsgBox "HIGH risk claimants highlighted in red.", vbInformation, "Done"
End Sub

'MACRO 3: For missed payments
Sub GenerateMissedPaymentsList()
    Dim srcWs As Worksheet, newWs As Worksheet
    Dim lastRow As Long, i As Long, newRow As Long
    Set srcWs = ThisWorkbook.Sheets("Payment Log")
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

    On Error Resume Next
    ThisWorkbook.Sheets("Missed Payments List").Delete
    On Error GoTo 0

    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = "Missed Payments List"

    newWs.Cells(1, 1).Value = "Missed Payment — Letter Generation List"
    newWs.Cells(1, 1).Font.Bold = True
    newWs.Cells(1, 1).Font.Size = 12

    Dim headers(1 To 5) As String
    headers(1) = "Claimant ID"
    headers(2) = "Claimant Name"
    headers(3) = "Month"
    headers(4) = "Expected Amount"
    headers(5) = "Action Required"

    For i = 1 To 5
        newWs.Cells(3, i).Value = headers(i)
        newWs.Cells(3, i).Font.Bold = True
        newWs.Cells(3, i).Interior.Color = RGB(46, 117, 182)
        newWs.Cells(3, i).Font.Color = RGB(255, 255, 255)
    Next i

    newRow = 4
    For i = 4 To lastRow
        If srcWs.Cells(i, 9).Value = "Missed" Then
            newWs.Cells(newRow, 1).Value = srcWs.Cells(i, 1).Value
            newWs.Cells(newRow, 2).Value = srcWs.Cells(i, 2).Value
            newWs.Cells(newRow, 3).Value = srcWs.Cells(i, 4).Value
            newWs.Cells(newRow, 4).Value = srcWs.Cells(i, 6).Value
            newWs.Cells(newRow, 5).Value = "Issue payment + send notification letter"
            newRow = newRow + 1
        End If
    Next i

    newWs.Columns("A:E").AutoFit
    MsgBox "Missed Payments list created (" & (newRow - 4) & " records).", vbInformation, "Done"
End Sub

'MACRO 4: Resolved discrepancy
Sub MarkResolved()
    Dim ws As Worksheet
    Dim selectedRow As Long
    Set ws = ThisWorkbook.Sheets("Discrepancy Register")

    If ActiveSheet.Name <> "Discrepancy Register" Then
        MsgBox "Please select a row in the Discrepancy Register sheet first.", vbExclamation
        Exit Sub
    End If

    selectedRow = ActiveCell.Row
    If selectedRow < 4 Then
        MsgBox "Please select a data row (row 4 or below).", vbExclamation
        Exit Sub
    End If

    ws.Cells(selectedRow, 11).Value = "Resolved"
    ws.Rows(selectedRow).Interior.Color = RGB(226, 239, 218)
    MsgBox "Row " & selectedRow & " marked as Resolved.", vbInformation, "Done"
End Sub

'MACRO 5: Export Discrepancy Register
Sub ExportDiscrepancyRegister()
    Dim srcWs As Worksheet
    Dim newWb As Workbook
    Set srcWs = ThisWorkbook.Sheets("Discrepancy Register")

    Set newWb = Workbooks.Add
    srcWs.UsedRange.Copy
    newWb.Sheets(1).Paste
    newWb.Sheets(1).Cells.EntireColumn.AutoFit

    Dim savePath As String
    savePath = Environ("USERPROFILE") & "\\Desktop\\Discrepancy_Review_" & Format(Now(), "YYYYMMDD") & ".xlsx"
    Application.DisplayAlerts = False
    newWb.SaveAs savePath, FileFormat:=51
    newWb.Close
    Application.DisplayAlerts = True

    MsgBox "Exported to: " & savePath, vbInformation, "Export Complete"
End Sub