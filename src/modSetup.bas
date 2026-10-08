Attribute VB_Name = "modSetup"
Option Explicit

' Creates sheets and navigation shapes; safe to rerun without deleting data.
Public Sub InitializeInvoiceApp()
    Dim dash As Worksheet
    On Error GoTo Failure
    EnsureDataSheet "Customers", Array("Customer ID", "Customer Name", "Email", "Country", "Status", "Created At")
    EnsureDataSheet "Products", Array("Product ID", "Product/Service", "Unit Price", "Demo Tax Rate", "Status", "Created At")
    EnsureDataSheet "Audit Log", Array("Timestamp", "Action", "Record ID", "Details")
    EnsureDataSheet "Quotations", Array("Quote ID", "Created At", "Customer ID", "Customer", "Discount Rate", "Net", "Discount", "Tax", "Gross", "Status")
    EnsureDataSheet "Quotation Lines", Array("Quote ID", "Product ID", "Description", "Quantity", "Unit Price", "Tax Rate", "Net", "Discount", "Tax", "Gross")
    SetupQuoteDraft
    Set dash = EnsureDashboard()
    BuildNavigation dash
    RefreshCatalogKpis
    dash.Activate
    MsgBox "Invoice & Quotation Automation v0.2 is ready.", vbInformation
    Exit Sub
Failure:
    MsgBox "Setup failed: " & Err.Description, vbExclamation
End Sub

Private Sub EnsureDataSheet(ByVal name As String, ByVal headers As Variant)
    Dim ws As Worksheet, c As Long, used As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = name
    End If
    If Len(Trim$(CStr(ws.Cells(1, 1).Value2))) = 0 Then
        For c = LBound(headers) To UBound(headers)
            ws.Cells(1, c + 1).Value2 = CStr(headers(c))
        Next c
    End If
    used = UBound(headers) + 1
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, used))
        .Interior.Color = RGB(25, 53, 85)
        .Font.Color = vbWhite
        .Font.Bold = True
        .RowHeight = 29
    End With
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 22
    If name = "Products" Then ws.Columns("C").NumberFormat = "#,##0.00"
    If name = "Customers" Or name = "Products" Then
        ws.Columns("A").NumberFormat = "@"
    End If
End Sub

Private Function EnsureDashboard() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = "Dashboard"
    End If
    Set EnsureDashboard = ws
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("E").ColumnWidth = 26
    With ws.Range("A1:E2")
        .Merge
        .Interior.Color = RGB(25, 53, 85)
        .Font.Color = vbWhite
        .Font.Name = "Calibri"
        .Font.Size = 19
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A1").Value2 = "INVOICE & QUOTATION AUTOMATION | v0.2"
    ws.Range("A4").Value2 = "Customer and product catalogs — Excel VBA"
    ws.Range("A6").Value2 = "Active customers"
    ws.Range("A7").Value2 = "Active products/services"
    ws.Range("A8").Value2 = "Last refresh"
    ws.Range("A6:A8").Font.Bold = True
    ws.Range("B6:B8").Interior.Color = RGB(234, 243, 249)
    ws.Range("A9").Value2 = "Last quotation"
    ws.Range("A12").Value2 = "v0.2: quotes are demo records, not legal invoices. PDF / Word / Outlook planned."
    ws.Range("A12:E13").Merge
    ws.Range("A12:E13").WrapText = True
    ws.Range("A12").Font.Color = RGB(90, 97, 110)
End Function

Private Sub BuildNavigation(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "inv_" Then ws.Shapes(i).Delete
    Next i
    DashboardButton ws, "Add customer", "AddCustomer", ws.Range("A10"), RGB(27, 126, 88)
    DashboardButton ws, "Add product/service", "AddProduct", ws.Range("B10"), RGB(27, 126, 88)
    DashboardButton ws, "Customers", "ShowCustomers", ws.Range("D6"), RGB(38, 105, 158)
    DashboardButton ws, "Products", "ShowProducts", ws.Range("E6"), RGB(38, 105, 158)
    DashboardButton ws, "Refresh KPI", "RefreshCatalogKpis", ws.Range("D8"), RGB(38, 105, 158)
    DashboardButton ws, "Quote Draft", "PrepareQuoteDraft", ws.Range("D10"), RGB(38, 105, 158)
    DashboardButton ws, "Create Quote", "CreateQuotation", ws.Range("E10"), RGB(27, 126, 88)
End Sub

Private Sub DashboardButton(ByVal ws As Worksheet, ByVal caption As String, _
                            ByVal macroName As String, ByVal anchor As Range, _
                            ByVal fillColor As Long)
    Dim shape As Shape
    Set shape = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, anchor.Width, 30)
    With shape
        .Name = "inv_" & macroName
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = caption
        .TextFrame.Characters.Font.Color = vbWhite
        .TextFrame.Characters.Font.Bold = True
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
        .Placement = xlMoveAndSize
    End With
End Sub

Private Sub SetupQuoteDraft()
    Dim ws As Worksheet
    Dim i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Quote Draft")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Quote Draft"
    End If
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 46
    ws.Columns("D").ColumnWidth = 25
    ws.Range("A1:D2").Merge
    With ws.Range("A1:D2")
        .Interior.Color = RGB(25, 53, 85)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 17
    End With
    ws.Range("A1").Value2 = "QUOTATION DRAFT | v0.2"
    ws.Range("A4").Value2 = "Customer ID"
    ws.Range("A5").Value2 = "Discount % (0-100)"
    ws.Range("A7").Value2 = "Add catalog product IDs and positive quantities below."
    ws.Range("A9").Value2 = "Product ID"
    ws.Range("B9").Value2 = "Quantity"
    ws.Range("A9:B9").Interior.Color = RGB(25, 53, 85)
    ws.Range("A9:B9").Font.Color = vbWhite
    ws.Range("A9:B9").Font.Bold = True
    ws.Range("A10:B29").Interior.Color = RGB(236, 245, 251)
    ws.Range("A10:A29").NumberFormat = "@"
    ws.Range("A31").Value2 = "Click Create Quote on the Dashboard or run CreateQuotation (Alt+F8)."
End Sub
