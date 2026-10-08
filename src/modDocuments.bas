Attribute VB_Name = "modDocuments"
Option Explicit

' v0.3 - Windows Word Desktop / Classic Outlook, late bound.
' Preview, requires manual Excel 2019 / Word 2019 / Outlook testing.
' NO automatic sending. All invoices are visibly DEMO / NOT FISCAL.

Public Sub ExportQuotationWordPdf()
    Dim quoteId As String, fileName As Variant
    Dim wordApp As Object, doc As Object
    Dim quoteRow As Long, q As Worksheet, lines As Worksheet
    Dim customer As String, created As String
    Dim outDocx As String, outPdf As String, basePath As String
    Dim wordTable As Object, atEnd As Object
    Dim r As Long, n As Long, last As Long, i As Long
    Dim savedDocx As Boolean, savedPdf As Boolean
    On Error GoTo Failed

    quoteId = Trim$(InputBox("Enter saved quote ID (example QUO-0001):", _
                             "Export quotation", CStr(ThisWorkbook.Worksheets("Dashboard").Range("B9").Value2)))
    If Len(quoteId) = 0 Then Exit Sub
    Set q = ThisWorkbook.Worksheets("Quotations")
    Set lines = ThisWorkbook.Worksheets("Quotation Lines")
    quoteRow = FindRowById(q, quoteId)
    If quoteRow = 0 Then Err.Raise vbObjectError + 410, , "Quote not found: " & quoteId
    customer = CStr(q.Cells(quoteRow, 4).Value2)

    fileName = Application.GetSaveAsFilename( _
             InitialFilename:=quoteId & "_Quotation.docx", _
             FileFilter:="Word Document (*.docx),*.docx", _
             Title:="Save demo quotation")
    If VarType(fileName) = vbBoolean Then Exit Sub
    outDocx = CStr(fileName)
    If LCase$(Right$(outDocx, 5)) <> ".docx" Then outDocx = outDocx & ".docx"
    outPdf = Left$(outDocx, Len(outDocx) - 5) & ".pdf"
    If Len(Dir$(outDocx)) <> 0 Or Len(Dir$(outPdf)) <> 0 Then
        Err.Raise vbObjectError + 411, , "Output DOCX or PDF already exists. Choose a new name."
    End If

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set doc = wordApp.Documents.Add
    doc.Content.InsertAfter "QUOTATION (DEMO)" & vbCrLf
    doc.Content.InsertAfter "Reference: " & quoteId & vbCrLf
    doc.Content.InsertAfter "Customer: " & customer & vbCrLf
    doc.Content.InsertAfter "Created: " & Format$(q.Cells(quoteRow, 2).Value, "yyyy-mm-dd") & vbCrLf
    doc.Content.InsertAfter "This is a sample commercial quotation, not a tax invoice." & vbCrLf & vbCrLf

    last = lines.Cells(lines.Rows.Count, 1).End(xlUp).Row
    For r = 2 To last
        If StrComp(CStr(lines.Cells(r, 1).Value2), quoteId, vbTextCompare) = 0 Then n = n + 1
    Next r
    If n = 0 Then Err.Raise vbObjectError + 412, , "Quote has no saved line items."

    Set atEnd = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    Set wordTable = doc.Tables.Add(atEnd, n + 1, 6)
    wordTable.Cell(1, 1).Range.Text = "Product"
    wordTable.Cell(1, 2).Range.Text = "Description"
    wordTable.Cell(1, 3).Range.Text = "Qty"
    wordTable.Cell(1, 4).Range.Text = "Unit"
    wordTable.Cell(1, 5).Range.Text = "Tax %"
    wordTable.Cell(1, 6).Range.Text = "Line total"
    i = 1
    For r = 2 To last
        If StrComp(CStr(lines.Cells(r, 1).Value2), quoteId, vbTextCompare) = 0 Then
            i = i + 1
            wordTable.Cell(i, 1).Range.Text = CStr(lines.Cells(r, 2).Value2)
            wordTable.Cell(i, 2).Range.Text = CStr(lines.Cells(r, 3).Value2)
            wordTable.Cell(i, 3).Range.Text = CStr(lines.Cells(r, 4).Value2)
            wordTable.Cell(i, 4).Range.Text = Format$(lines.Cells(r, 5).Value2, "0.00")
            wordTable.Cell(i, 5).Range.Text = Format$(CDbl(lines.Cells(r, 6).Value2) * 100#, "0.##")
            wordTable.Cell(i, 6).Range.Text = Format$(lines.Cells(r, 10).Value2, "0.00")
        End If
    Next r
    wordTable.Rows(1).Range.Bold = True
    wordTable.Borders.Enable = True

    Set atEnd = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    atEnd.InsertAfter vbCrLf & _
        "Subtotal: " & Format$(q.Cells(quoteRow, 6).Value2, "0.00") & vbCrLf & _
        "Discount: " & Format$(q.Cells(quoteRow, 7).Value2, "0.00") & vbCrLf & _
        "Demo tax: " & Format$(q.Cells(quoteRow, 8).Value2, "0.00") & vbCrLf & _
        "TOTAL: " & Format$(q.Cells(quoteRow, 9).Value2, "0.00") & vbCrLf & _
        "Currency: not configured (demo amounts only)."

    doc.SaveAs2 FileName:=outDocx, FileFormat:=16
    savedDocx = True
    doc.ExportAsFixedFormat OutputFileName:=outPdf, ExportFormat:=17
    savedPdf = True
    doc.Close False
    Set doc = Nothing
    wordApp.Quit
    Set wordApp = Nothing
    WriteAudit "EXPORT_QUOTE_DOCS", quoteId, "DOCX and PDF generated"
    MsgBox "Quotation exported:" & vbCrLf & outDocx & vbCrLf & outPdf, vbInformation
    Exit Sub
Failed:
    Dim problem As String
    problem = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    On Error GoTo 0
    MsgBox "Export failed: " & problem & vbCrLf & _
           "Any previously saved files are left in place; inspect before retrying.", vbExclamation
End Sub

Public Sub CreateDemoInvoice()
    Dim quoteId As String, q As Worksheet, invoices As Worksheet
    Dim qr As Long, r As Long, id As String, n As Long
    On Error GoTo Failed
    quoteId = Trim$(InputBox("Saved quote ID:", "Create DEMO invoice", _
                    CStr(ThisWorkbook.Worksheets("Dashboard").Range("B9").Value2)))
    If Len(quoteId) = 0 Then Exit Sub
    Set q = ThisWorkbook.Worksheets("Quotations")
    Set invoices = ThisWorkbook.Worksheets("Demo Invoices")
    qr = FindRowById(q, quoteId)
    If qr = 0 Then Err.Raise vbObjectError + 420, , "Unknown quote: " & quoteId
    For r = 2 To invoices.Cells(invoices.Rows.Count, 1).End(xlUp).Row
        If StrComp(CStr(invoices.Cells(r, 3).Value2), quoteId, vbTextCompare) = 0 Then _
            Err.Raise vbObjectError + 421, , "This quotation already has a demo invoice."
    Next r
    n = invoices.Cells(invoices.Rows.Count, 1).End(xlUp).Row
    id = "DEMO-INV-" & Format$(n, "0000")
    r = n + 1
    invoices.Cells(r, 1).NumberFormat = "@"
    invoices.Cells(r, 1).Value2 = id
    invoices.Cells(r, 2).Value2 = Now
    invoices.Cells(r, 3).Value2 = quoteId
    invoices.Cells(r, 4).Value2 = q.Cells(qr, 3).Value2
    invoices.Cells(r, 5).Value2 = q.Cells(qr, 4).Value2
    invoices.Cells(r, 6).Value2 = q.Cells(qr, 9).Value2
    invoices.Cells(r, 7).Value2 = Date + 30
    invoices.Cells(r, 8).Value2 = "DEMO - NOT FISCAL"
    WriteAudit "CREATE_DEMO_INVOICE", id, quoteId
    MsgBox "Created " & id & " (demonstration record only, NOT a fiscal invoice).", vbInformation
    Exit Sub
Failed:
    MsgBox "Demo invoice creation failed: " & Err.Description, vbExclamation
End Sub

Public Sub PrepareQuotationOutlookDraft()
    Dim quoteId As String, q As Worksheet, customers As Worksheet
    Dim qr As Long, cr As Long, r As Long, custId As String, email As String
    Dim picked As Variant, outlookApp As Object, mail As Object
    On Error GoTo Failed
    quoteId = Trim$(InputBox("Quote ID for the email subject:", "Outlook draft", _
                    CStr(ThisWorkbook.Worksheets("Dashboard").Range("B9").Value2)))
    If Len(quoteId) = 0 Then Exit Sub
    Set q = ThisWorkbook.Worksheets("Quotations")
    Set customers = ThisWorkbook.Worksheets("Customers")
    qr = FindRowById(q, quoteId)
    If qr = 0 Then Err.Raise vbObjectError + 430, , "Unknown quote."
    custId = CStr(q.Cells(qr, 3).Value2)
    cr = FindRowById(customers, custId)
    If cr > 0 Then email = Trim$(CStr(customers.Cells(cr, 3).Value2))
    picked = Application.GetOpenFilename("PDF files (*.pdf),*.pdf", , _
                                         "Select the generated quotation PDF")
    If VarType(picked) = vbBoolean Then Exit Sub
    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)
    mail.To = email
    mail.Subject = "Quotation " & quoteId
    mail.Body = "Hello," & vbCrLf & vbCrLf & _
         "Please find attached our quotation for your review." & vbCrLf & _
         "This is a draft message. Please check the attachment and recipient before sending." & vbCrLf
    mail.Attachments.Add CStr(picked)
    mail.Display
    ' Deliberately NO .Send; recipient must review the message.
    WriteAudit "PREPARE_OUTLOOK_DRAFT", quoteId, "Manual review required"
    MsgBox "Outlook draft opened. Review the recipient and attachment; it was NOT sent.", vbInformation
    Exit Sub
Failed:
    MsgBox "Could not prepare Outlook draft. Classic Outlook Desktop is required." & vbCrLf & _
           Err.Description, vbExclamation
End Sub

Private Function FindRowById(ByVal ws As Worksheet, ByVal id As String) As Long
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value2)), id, vbTextCompare) = 0 Then
            FindRowById = r
            Exit Function
        End If
    Next r
End Function
