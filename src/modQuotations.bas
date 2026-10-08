Attribute VB_Name = "modQuotations"
Option Explicit

' v0.2 quotation MVP. Demonstration only; not a legally compliant invoice.
' Unit prices/tax rates snapshotted at creation; all amounts use 2 decimals.
Private Const FIRST_LINE As Long = 10
Private Const LAST_LINE As Long = 29

Public Sub PrepareQuoteDraft()
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("Quote Draft")
    ws.Activate
    MsgBox "Enter Customer ID in B4, optional discount in B5 (0 to 100)," & vbCrLf & _
           "then Product IDs in A10:A29 and quantities in B10:B29." & vbCrLf & _
           "Run CreateQuotation when ready.", vbInformation, "Quote Draft"
    Exit Sub
Failed:
    MsgBox "Run InitializeInvoiceApp first: " & Err.Description, vbExclamation
End Sub

Public Sub CreateQuotation()
    Dim draft As Worksheet, customers As Worksheet, products As Worksheet
    Dim quotes As Worksheet, lines As Worksheet, audit As Worksheet
    Dim customerId As String, customerName As String, quoteId As String
    Dim rawDiscount As Variant, discountRate As Double, productId As String
    Dim productRow As Long, customerRow As Long, i As Long, count As Long
    Dim qty As Double, price As Currency, taxRate As Double
    Dim net As Currency, discountValue As Currency, taxable As Currency, taxAmount As Currency
    Dim totalNet As Currency, totalDiscount As Currency, totalTax As Currency, totalGross As Currency
    Dim staged As Collection, one As Variant, item(1 To 8) As Variant
    Dim quoteRow As Long, lineRow As Long, nextNumber As Long
    Dim hadError As Boolean, errText As String
    On Error GoTo Failed

    Set draft = ThisWorkbook.Worksheets("Quote Draft")
    Set customers = ThisWorkbook.Worksheets("Customers")
    Set products = ThisWorkbook.Worksheets("Products")
    Set quotes = ThisWorkbook.Worksheets("Quotations")
    Set lines = ThisWorkbook.Worksheets("Quotation Lines")

    customerId = Trim$(CStr(draft.Range("B4").Value2))
    If Len(customerId) = 0 Then Err.Raise vbObjectError + 310, , "Customer ID (B4) is required."
    customerRow = FindActiveId(customers, customerId)
    If customerRow = 0 Then Err.Raise vbObjectError + 311, , "Unknown or inactive customer ID: " & customerId
    customerName = CStr(customers.Cells(customerRow, 2).Value2)

    rawDiscount = draft.Range("B5").Value2
    If Len(Trim$(CStr(rawDiscount))) = 0 Then rawDiscount = 0
    If Not IsNumeric(rawDiscount) Then Err.Raise vbObjectError + 312, , "Discount must be numeric."
    discountRate = CDbl(rawDiscount)
    If discountRate < 0 Or discountRate > 100 Then _
        Err.Raise vbObjectError + 313, , "Discount must be between 0 and 100."

    Set staged = New Collection
    For i = FIRST_LINE To LAST_LINE
        productId = Trim$(CStr(draft.Cells(i, 1).Value2))
        If Len(productId) > 0 Or Len(Trim$(CStr(draft.Cells(i, 2).Value2))) > 0 Then
            If Len(productId) = 0 Then _
                Err.Raise vbObjectError + 314, , "Missing product ID at row " & i
            productRow = FindActiveId(products, productId)
            If productRow = 0 Then _
                Err.Raise vbObjectError + 315, , "Unknown/inactive product at row " & i & ": " & productId
            If Not IsNumeric(draft.Cells(i, 2).Value2) Then _
                Err.Raise vbObjectError + 316, , "Invalid quantity at row " & i
            qty = CDbl(draft.Cells(i, 2).Value2)
            If qty <= 0 Or qty > 100000 Then _
                Err.Raise vbObjectError + 317, , "Quantity must be positive and <=100000 at row " & i
            price = CCur(products.Cells(productRow, 3).Value2)
            taxRate = CDbl(products.Cells(productRow, 4).Value2)
            If price < 0 Or taxRate < 0 Or taxRate > 1 Then _
                Err.Raise vbObjectError + 318, , "Product price/tax configuration is invalid."
            net = RoundMoney(CDbl(price) * qty)
            discountValue = RoundMoney(CDbl(net) * discountRate / 100#)
            taxable = net - discountValue
            taxAmount = RoundMoney(CDbl(taxable) * taxRate)
            count = count + 1
            item(1) = productId
            item(2) = CStr(products.Cells(productRow, 2).Value2)
            item(3) = qty
            item(4) = price
            item(5) = taxRate
            item(6) = net
            item(7) = discountValue
            item(8) = taxAmount
            staged.Add Array(item(1), item(2), item(3), item(4), item(5), item(6), item(7), item(8))
            totalNet = totalNet + net
            totalDiscount = totalDiscount + discountValue
            totalTax = totalTax + taxAmount
        End If
    Next i
    If count = 0 Then Err.Raise vbObjectError + 319, , "Enter at least one product line."
    totalGross = totalNet - totalDiscount + totalTax

    nextNumber = NextQuoteNumber(quotes)
    quoteId = "QUO-" & Format$(nextNumber, "0000")
    quoteRow = quotes.Cells(quotes.Rows.Count, 1).End(xlUp).Row + 1
    lineRow = lines.Cells(lines.Rows.Count, 1).End(xlUp).Row + 1
    If quoteRow + 1 > quotes.Rows.Count Or lineRow + count >= lines.Rows.Count Then _
        Err.Raise vbObjectError + 320, , "Quotation storage is full."

    ' Validate everything before the first persistent write.
    quotes.Cells(quoteRow, 1).NumberFormat = "@"
    quotes.Cells(quoteRow, 1).Value2 = quoteId
    quotes.Cells(quoteRow, 2).Value2 = Now
    quotes.Cells(quoteRow, 3).Value2 = customerId
    quotes.Cells(quoteRow, 4).Value2 = customerName
    quotes.Cells(quoteRow, 5).Value2 = discountRate / 100#
    quotes.Cells(quoteRow, 6).Value2 = totalNet
    quotes.Cells(quoteRow, 7).Value2 = totalDiscount
    quotes.Cells(quoteRow, 8).Value2 = totalTax
    quotes.Cells(quoteRow, 9).Value2 = totalGross
    quotes.Cells(quoteRow, 10).Value2 = "Draft"

    For Each one In staged
        lines.Cells(lineRow, 1).NumberFormat = "@"
        lines.Cells(lineRow, 1).Value2 = quoteId
        lines.Cells(lineRow, 2).Value2 = CStr(one(0))
        lines.Cells(lineRow, 3).Value2 = CStr(one(1))
        lines.Cells(lineRow, 4).Value2 = CDbl(one(2))
        lines.Cells(lineRow, 5).Value2 = CCur(one(3))
        lines.Cells(lineRow, 6).Value2 = CDbl(one(4))
        lines.Cells(lineRow, 7).Value2 = CCur(one(5))
        lines.Cells(lineRow, 8).Value2 = CCur(one(6))
        lines.Cells(lineRow, 9).Value2 = CCur(one(7))
        lines.Cells(lineRow, 10).Value2 = CCur(one(5)) - CCur(one(6)) + CCur(one(7))
        lineRow = lineRow + 1
    Next one
    WriteAudit "CREATE_QUOTATION", quoteId, customerId & " / " & CStr(count) & " lines"
    ThisWorkbook.Worksheets("Dashboard").Range("B9").Value2 = quoteId
    MsgBox "Quotation " & quoteId & " created." & vbCrLf & _
           "Gross total: " & Format$(totalGross, "0.00") & vbCrLf & _
           "Demo only: no invoice/PDF generated.", vbInformation
    quotes.Activate
    Exit Sub

Failed:
    errText = Err.Description
    MsgBox "Quotation was not completed: " & errText & vbCrLf & _
           "Check the Quotations and Quotation Lines sheets before retrying if an Excel write failed.", vbExclamation
End Sub

Private Function FindActiveId(ByVal ws As Worksheet, ByVal id As String) As Long
    Dim i As Long, last As Long
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To last
        If StrComp(Trim$(CStr(ws.Cells(i, 1).Value2)), id, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(i, 5).Value2)), "Active", vbTextCompare) = 0 Then
                FindActiveId = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function NextQuoteNumber(ByVal ws As Worksheet) As Long
    Dim i As Long, id As String, number As Long
    NextQuoteNumber = 1
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        id = CStr(ws.Cells(i, 1).Value2)
        If Left$(id, 4) = "QUO-" Then
            If IsNumeric(Mid$(id, 5)) Then
                number = CLng(Mid$(id, 5))
                If number >= NextQuoteNumber Then NextQuoteNumber = number + 1
            End If
        End If
    Next i
End Function

Private Function RoundMoney(ByVal amount As Double) As Currency
    ' Excel ROUND, not VBA Round (banker's rounding).
    RoundMoney = CCur(Application.WorksheetFunction.Round(amount, 2))
End Function
