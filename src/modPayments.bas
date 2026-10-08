Attribute VB_Name = "modPayments"
Option Explicit

' v0.4 DEMONSTRATION payment tracking (NOT real accounting/fiscal software).
' Recomputes paid, balance and effective status from append-only Payments rows.
' Required sheets: Demo Invoices, Payments, Dashboard, Audit Log.
Private Const MONEY_EPSILON As Double = 0.005

Public Sub RecordDemoPayment()
    Dim invoices As Worksheet, payments As Worksheet
    Dim invoiceId As String, rawAmount As String, note As String
    Dim invoiceRow As Long, target As Currency, paid As Currency
    Dim amount As Currency, pending As Currency, nextRow As Long, paymentId As String
    Dim answer As VbMsgBoxResult
    On Error GoTo Failed
    Set invoices = ThisWorkbook.Worksheets("Demo Invoices")
    Set payments = ThisWorkbook.Worksheets("Payments")
    invoiceId = Trim$(InputBox("Demo invoice ID (example DEMO-INV-0001):", _
                             "Record demo payment"))
    If Len(invoiceId) = 0 Then Exit Sub
    invoiceRow = FindInvoiceRow(invoices, invoiceId)
    If invoiceRow = 0 Then Err.Raise vbObjectError + 501, , "Unknown invoice ID."
    target = CCur(invoices.Cells(invoiceRow, 6).Value2)
    If target <= 0 Then Err.Raise vbObjectError + 502, , "Invoice amount must be positive."
    paid = PaidForInvoice(payments, invoiceId)
    pending = target - paid
    If pending <= 0 Then Err.Raise vbObjectError + 503, , "Invoice is fully paid."
    rawAmount = Trim$(InputBox("Amount to record (available: " & _
                        Format$(pending, "0.00") & "). Use your Windows decimal separator:", _
                        "Record demo payment"))
    If Len(rawAmount) = 0 Then Exit Sub
    If Not IsNumeric(rawAmount) Then Err.Raise vbObjectError + 504, , "Amount must be numeric."
    If CDbl(rawAmount) <= 0 Then Err.Raise vbObjectError + 505, , "Amount must be positive."
    amount = RoundCurrency(CDbl(rawAmount))
    If amount <= 0 Then Err.Raise vbObjectError + 506, , "Amount rounds to zero."
    If amount > pending Then Err.Raise vbObjectError + 507, , _
                      "Payment exceeds remaining balance: " & Format$(pending, "0.00")
    note = Trim$(InputBox("Optional reference/note (fictional data only):", _
                          "Payment reference"))
    answer = MsgBox("Record DEMO payment " & Format$(amount, "0.00") & _
             " against " & invoiceId & "?" & vbCrLf & _
             "This is an append-only record, not a real bank transaction.", _
             vbYesNo + vbQuestion, "Confirm demo payment")
    If answer <> vbYes Then Exit Sub
    nextRow = NextDataRow(payments)
    paymentId = "PAY-" & Format$(NextPaymentNumber(payments), "0000")
    payments.Cells(nextRow, 1).NumberFormat = "@"
    payments.Cells(nextRow, 1).Value2 = paymentId
    payments.Cells(nextRow, 2).Value2 = Now
    payments.Cells(nextRow, 3).NumberFormat = "@"
    payments.Cells(nextRow, 3).Value2 = invoiceId
    payments.Cells(nextRow, 4).Value2 = amount
    payments.Cells(nextRow, 5).Value2 = note
    payments.Cells(nextRow, 4).NumberFormat = "#,##0.00"
    WriteAudit "RECORD_DEMO_PAYMENT", paymentId, invoiceId & " / " & _
               Format$(amount, "0.00")
    RefreshPaymentDashboard
    MsgBox "Recorded demo payment " & paymentId & ".", vbInformation
    Exit Sub
Failed:
    MsgBox "Payment was NOT confirmed: " & Err.Description, vbExclamation
End Sub

Public Sub RefreshPaymentDashboard()
    Dim invoices As Worksheet, payments As Worksheet, dash As Worksheet
    Dim r As Long, last As Long, invoiceId As String
    Dim total As Currency, paid As Currency, remaining As Currency
    Dim allTotal As Currency, allPaid As Currency, allRemaining As Currency
    Dim overdueTotal As Currency, countInvoices As Long
    Dim dueDate As Variant, status As String
    On Error GoTo Failed
    Set invoices = ThisWorkbook.Worksheets("Demo Invoices")
    Set payments = ThisWorkbook.Worksheets("Payments")
    Set dash = ThisWorkbook.Worksheets("Dashboard")
    last = invoices.Cells(invoices.Rows.Count, 1).End(xlUp).Row
    For r = 2 To last
        invoiceId = Trim$(CStr(invoices.Cells(r, 1).Value2))
        If Len(invoiceId) > 0 Then
            If Not IsNumeric(invoices.Cells(r, 6).Value2) Then _
                Err.Raise vbObjectError + 511, , "Invalid invoice gross at row " & r
            total = CCur(invoices.Cells(r, 6).Value2)
            If total < 0 Then Err.Raise vbObjectError + 512, , "Negative invoice gross at row " & r
            paid = PaidForInvoice(payments, invoiceId)
            If paid > total Then Err.Raise vbObjectError + 513, , _
                "Recorded payments exceed invoice " & invoiceId
            remaining = total - paid
            dueDate = invoices.Cells(r, 7).Value
            If remaining = 0 Then
                status = "Paid"
            ElseIf IsDate(dueDate) Then
                If CDate(dueDate) < Date Then
                    status = "Overdue"
                ElseIf paid > 0 Then
                    status = "Partially Paid"
                Else
                    status = "Unpaid"
                End If
            ElseIf paid > 0 Then
                status = "Partially Paid"
            Else
                status = "Unpaid"
            End If
            invoices.Cells(r, 9).Value2 = paid
            invoices.Cells(r, 10).Value2 = remaining
            invoices.Cells(r, 11).Value2 = status
            countInvoices = countInvoices + 1
            allTotal = allTotal + total
            allPaid = allPaid + paid
            allRemaining = allRemaining + remaining
            If status = "Overdue" Then overdueTotal = overdueTotal + remaining
        End If
    Next r
    dash.Range("A19").Value2 = "DEMO PAYMENT OVERVIEW"
    dash.Range("A20").Value2 = "Demo invoices"
    dash.Range("B20").Value2 = countInvoices
    dash.Range("A21").Value2 = "Total demo billed"
    dash.Range("B21").Value2 = allTotal
    dash.Range("A22").Value2 = "Recorded demo payments"
    dash.Range("B22").Value2 = allPaid
    dash.Range("A23").Value2 = "Outstanding balance"
    dash.Range("B23").Value2 = allRemaining
    dash.Range("A24").Value2 = "Overdue balance"
    dash.Range("B24").Value2 = overdueTotal
    dash.Range("B21:B24").NumberFormat = "#,##0.00"
    With dash.Range("A19:B19")
        .Interior.Color = RGB(25, 53, 85)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    dash.Range("A20:A24").Font.Bold = True
    dash.Range("B20:B24").Interior.Color = RGB(234, 243, 249)
    Exit Sub
Failed:
    MsgBox "Unable to refresh payment data: " & Err.Description, vbExclamation
End Sub

Public Sub FilterDemoInvoices()
    Dim ws As Worksheet, choice As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("Demo Invoices")
    choice = Trim$(InputBox( _
       "Enter All, Paid, Unpaid, Partially Paid or Overdue:", _
       "Filter demo invoices", "All"))
    If Len(choice) = 0 Then Exit Sub
    RefreshPaymentDashboard
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Select Case UCase$(choice)
        Case "ALL"
        Case "PAID", "UNPAID", "PARTIALLY PAID", "OVERDUE"
            ws.Range("A1:K" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row) _
              .AutoFilter Field:=11, Criteria1:=choice
        Case Else
            Err.Raise vbObjectError + 515, , "Unsupported filter selection."
    End Select
    ws.Activate
    Exit Sub
Failed:
    MsgBox "Unable to filter: " & Err.Description, vbExclamation
End Sub

Private Function FindInvoiceRow(ByVal ws As Worksheet, ByVal invoiceId As String) As Long
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value2)), invoiceId, vbTextCompare) = 0 Then
            FindInvoiceRow = r
            Exit Function
        End If
    Next r
End Function

Private Function PaidForInvoice(ByVal payments As Worksheet, ByVal invoiceId As String) As Currency
    Dim r As Long, value As Variant, paid As Currency
    For r = 2 To NextDataRow(payments) - 1
        If StrComp(Trim$(CStr(payments.Cells(r, 3).Value2)), invoiceId, vbTextCompare) = 0 Then
            value = payments.Cells(r, 4).Value2
            If Not IsNumeric(value) Then Err.Raise vbObjectError + 520, , _
                                        "Invalid payment amount in row " & r
            If CDbl(value) <= 0 Then Err.Raise vbObjectError + 521, , _
                                        "Nonpositive payment amount in row " & r
            paid = paid + CCur(value)
        End If
    Next r
    PaidForInvoice = paid
End Function

Private Function NextDataRow(ByVal ws As Worksheet) As Long
    NextDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If NextDataRow < 2 Then NextDataRow = 2
End Function

Private Function NextPaymentNumber(ByVal ws As Worksheet) As Long
    Dim r As Long, value As String, number As Long
    NextPaymentNumber = 1
    For r = 2 To NextDataRow(ws) - 1
        value = CStr(ws.Cells(r, 1).Value2)
        If Left$(value, 4) = "PAY-" Then
            If IsNumeric(Mid$(value, 5)) Then
                number = CLng(Mid$(value, 5))
                If number >= NextPaymentNumber Then NextPaymentNumber = number + 1
            End If
        End If
    Next r
End Function

Private Function RoundCurrency(ByVal amount As Double) As Currency
    RoundCurrency = CCur(Application.WorksheetFunction.Round(amount, 2))
End Function
