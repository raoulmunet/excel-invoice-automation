Attribute VB_Name = "modUI"
Option Explicit

' v1.0 visual layer. Does NOT clear data cells, tables, quotes, invoices or payments.
' Keep KPI storage references unchanged for backward compatibility.
Public Sub ApplyProfessionalTheme()
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Application.ScreenUpdating = False
    With ws
        .Activate
        .Cells.Font.Name = "Aptos"
        .Cells.Font.Size = 11
        .Range("A1:J27").RowHeight = 24
        .Rows("1:2").RowHeight = 29
        .Columns("A").ColumnWidth = 27
        .Columns("B").ColumnWidth = 22
        .Columns("C").ColumnWidth = 4
        .Columns("D").ColumnWidth = 24
        .Columns("E").ColumnWidth = 24
        .Columns("F").ColumnWidth = 3
        .Columns("G").ColumnWidth = 23
        .Columns("H").ColumnWidth = 23
        .Columns("I").ColumnWidth = 23
        .Columns("J").ColumnWidth = 23
        .Range("A1:J28").Interior.Color = RGB(248, 250, 253)
        .Range("A1:J2").UnMerge
        .Range("A1:J2").Merge
        With .Range("A1:J2")
            .Interior.Color = RGB(21, 43, 69)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 20
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
        End With
        .Range("A1").Value2 = "OFFICE AUTOMATION  |  BUSINESS CONTROL CENTER"
        .Range("A3").Value2 = "Excel + Word + PDF + Classic Outlook"
        .Range("A3").Font.Color = RGB(94, 108, 127)
        .Range("A4").Value2 = "LIVE OVERVIEW"
        SectionTitle ws, "A5:B5", "CUSTOMERS & CATALOG"
        SectionTitle ws, "A18:B18", "DEMO FINANCIAL OVERVIEW"
        SectionTitle ws, "G5:J5", "QUICK ACTIONS"
        SectionTitle ws, "G12:J12", "COMMERCIAL WORKFLOW"
        SectionTitle ws, "G18:J18", "PAYMENTS & ANALYTICS"
        .Range("A6:B9").Interior.Color = RGB(237, 244, 250)
        .Range("A20:B24").Interior.Color = RGB(237, 244, 250)
        .Range("A6:A9").Font.Color = RGB(53, 70, 92)
        .Range("A20:A24").Font.Color = RGB(53, 70, 92)
        .Range("A6:A9").Font.Bold = True
        .Range("A20:A24").Font.Bold = True
        .Range("B6:B9").Font.Color = RGB(18, 75, 120)
        .Range("B20:B24").Font.Color = RGB(18, 75, 120)
        .Range("B6:B9").Font.Bold = True
        .Range("B20:B24").Font.Bold = True
        .Range("B20:B24").NumberFormat = "#,##0.00"
        .Range("B20").NumberFormat = "0"
        .Range("A11:E14").UnMerge
        .Range("A11:E14").Merge
        .Range("A11").Value2 = "DEMONSTRATION ONLY  |  All quotations, invoices and payments are fictitious. These are not fiscal documents or live banking records. Review Outlook messages before manually sending."
        With .Range("A11:E14")
            .Interior.Color = RGB(255, 245, 229)
            .Font.Color = RGB(124, 79, 25)
            .Font.Size = 10
            .WrapText = True
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
        End With
        .Range("G25:J27").UnMerge
        .Range("G25:J27").Merge
        .Range("G25").Value2 = "WORKFLOW   Catalog  >  Quotation  >  DOCX/PDF  >  Demo Invoice  >  Payment  >  Outlook Draft"
        .Range("G25:J27").Interior.Color = RGB(231, 241, 249)
        .Range("G25:J27").Font.Color = RGB(44, 75, 103)
        .Range("G25:J27").WrapText = True
        .Range("G25:J27").VerticalAlignment = xlCenter
        .Range("G25:J27").HorizontalAlignment = xlCenter
        .Range("A27").Value2 = "v1.0 interface preview | Office 2019"
        .Range("A27").Font.Color = RGB(112, 125, 139)
    End With
    PlaceAction ws, "AddCustomer", "G7", 23
    PlaceAction ws, "AddProduct", "I7", 23
    PlaceAction ws, "ShowCustomers", "G9", 23
    PlaceAction ws, "ShowProducts", "I9", 23
    PlaceAction ws, "PrepareQuoteDraft", "G14", 23
    PlaceAction ws, "CreateQuotation", "I14", 23
    PlaceAction ws, "ExportQuotationWordPdf", "G16", 23
    PlaceAction ws, "CreateDemoInvoice", "I16", 23
    PlaceAction ws, "RecordDemoPayment", "G20", 23
    PlaceAction ws, "RefreshPaymentDashboard", "I20", 23
    PlaceAction ws, "FilterDemoInvoices", "G22", 23
    PlaceAction ws, "PrepareQuotationOutlookDraft", "I22", 23
    PlaceAction ws, "RefreshCatalogKpis", "D7", 22
    FormatDataSheets
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox "Dashboard styling failed: " & Err.Description, vbExclamation
End Sub

Private Sub SectionTitle(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String)
    With ws.Range(address)
        .UnMerge
        .Merge
        .Value2 = caption
        .Interior.Color = RGB(42, 82, 122)
        .Font.Color = vbWhite
        .Font.Size = 10
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
End Sub

Private Sub PlaceAction(ByVal ws As Worksheet, ByVal macroName As String, _
                        ByVal cell As String, ByVal width As Double)
    Dim sh As Shape, anchor As Range
    On Error Resume Next
    Set sh = ws.Shapes("inv_" & macroName)
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub
    Set anchor = ws.Range(cell)
    sh.Left = anchor.Left + 3
    sh.Top = anchor.Top + 2
    If width > 0 Then sh.Width = anchor.Width * 1.65
    sh.Height = 28
    sh.Placement = xlMove
    sh.Line.Visible = msoFalse
    sh.Fill.ForeColor.RGB = RGB(35, 111, 161)
    If macroName = "CreateQuotation" Or macroName = "RecordDemoPayment" Then _
        sh.Fill.ForeColor.RGB = RGB(25, 135, 99)
    With sh.TextFrame
        .Characters.Font.Name = "Calibri"
        .Characters.Font.Size = 10
        .Characters.Font.Bold = True
        .Characters.Font.Color = vbWhite
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
End Sub

Private Sub FormatDataSheets()
    Dim names As Variant, i As Long, ws As Worksheet, last As Long
    names = Array("Customers", "Products", "Quotations", "Quotation Lines", _
                  "Demo Invoices", "Payments", "Audit Log")
    For i = LBound(names) To UBound(names)
        Set ws = ThisWorkbook.Worksheets(CStr(names(i)))
        last = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If last >= 1 Then
            With ws.Range(ws.Cells(1, 1), ws.Cells(1, last))
                .Interior.Color = RGB(21, 43, 69)
                .Font.Color = vbWhite
                .Font.Bold = True
                .RowHeight = 30
                .VerticalAlignment = xlCenter
            End With
            ws.Rows("2:2").RowHeight = 22
            ws.Tab.Color = RGB(42, 82, 122)
        End If
    Next i
    ThisWorkbook.Worksheets("Quote Draft").Tab.Color = RGB(25, 135, 99)
    ThisWorkbook.Worksheets("Dashboard").Tab.Color = RGB(25, 135, 99)
End Sub
