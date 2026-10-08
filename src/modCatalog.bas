Attribute VB_Name = "modCatalog"
Option Explicit

' Invoice & Quotation Automation v0.1
' Source of truth: Customers and Products worksheets in ThisWorkbook.
' Test target: Windows 11 + Excel 2019 Desktop (runtime verification pending).
Private Const CUSTOMERS As String = "Customers"
Private Const PRODUCTS As String = "Products"

Public Sub AddCustomer()
    Dim name As String, email As String, country As String
    Dim ws As Worksheet, r As Long, customerId As String
    On Error GoTo Failure
    Set ws = RequireSheet(CUSTOMERS)
    name = Trim$(InputBox("Customer / business name:", "New customer"))
    If Len(name) = 0 Then Exit Sub
    email = Trim$(InputBox("Contact email (optional):", "New customer"))
    country = Trim$(InputBox("Country (optional):", "New customer"))
    If Len(email) > 0 And (InStr(1, email, "@") < 2 Or InStrRev(email, ".") < InStr(1, email, "@") + 2) Then
        MsgBox "Email does not look valid. Customer was not saved.", vbExclamation
        Exit Sub
    End If
    r = NextRow(ws)
    customerId = "CUS-" & Format$(NextSequence(ws, "CUS-"), "0000")
    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 1).Value2 = customerId
    ws.Cells(r, 2).Value2 = name
    ws.Cells(r, 3).Value2 = email
    ws.Cells(r, 4).Value2 = country
    ws.Cells(r, 5).Value2 = "Active"
    ws.Cells(r, 6).Value2 = Now
    ws.Cells(r, 6).NumberFormat = "yyyy-mm-dd hh:mm"
    WriteAudit "CREATE_CUSTOMER", customerId, name
    MsgBox "Customer saved: " & customerId, vbInformation
    Exit Sub
Failure:
    MsgBox "Unable to create customer: " & Err.Description, vbExclamation
End Sub

Public Sub AddProduct()
    Dim productName As String, rawPrice As String, rawTax As String
    Dim price As Double, tax As Double, ws As Worksheet, r As Long, productId As String
    On Error GoTo Failure
    Set ws = RequireSheet(PRODUCTS)
    productName = Trim$(InputBox("Product or service name:", "New product"))
    If Len(productName) = 0 Then Exit Sub
    rawPrice = Trim$(InputBox("Unit price (use your system's decimal separator):", "Unit price"))
    If Len(rawPrice) = 0 Then Exit Sub
    If Not IsNumeric(rawPrice) Then Err.Raise vbObjectError + 205, , "Price must be numeric."
    price = CDbl(rawPrice)
    If price < 0 Then Err.Raise vbObjectError + 206, , "Price cannot be negative."
    rawTax = Trim$(InputBox("Sample tax rate (%). Demo only; not legal tax advice:", "Tax rate", "0"))
    If Len(rawTax) = 0 Then Exit Sub
    If Not IsNumeric(rawTax) Then Err.Raise vbObjectError + 207, , "Tax rate must be numeric."
    tax = CDbl(rawTax)
    If tax < 0 Or tax > 100 Then Err.Raise vbObjectError + 208, , "Tax rate must be 0-100."
    r = NextRow(ws)
    productId = "PRD-" & Format$(NextSequence(ws, "PRD-"), "0000")
    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 1).Value2 = productId
    ws.Cells(r, 2).Value2 = productName
    ws.Cells(r, 3).Value2 = price
    ws.Cells(r, 4).Value2 = tax / 100#
    ws.Cells(r, 4).NumberFormat = "0.00%"
    ws.Cells(r, 5).Value2 = "Active"
    ws.Cells(r, 6).Value2 = Now
    ws.Cells(r, 6).NumberFormat = "yyyy-mm-dd hh:mm"
    WriteAudit "CREATE_PRODUCT", productId, productName
    MsgBox "Product saved: " & productId, vbInformation
    Exit Sub
Failure:
    MsgBox "Unable to create product: " & Err.Description, vbExclamation
End Sub

Public Sub ShowCustomers()
    On Error GoTo Failure
    RequireSheet(CUSTOMERS).Activate
    Exit Sub
Failure:
    MsgBox "Customers sheet not found. Run InitializeInvoiceApp first.", vbExclamation
End Sub

Public Sub ShowProducts()
    On Error GoTo Failure
    RequireSheet(PRODUCTS).Activate
    Exit Sub
Failure:
    MsgBox "Products sheet not found. Run InitializeInvoiceApp first.", vbExclamation
End Sub

Public Sub ShowDashboard()
    On Error GoTo Failure
    RequireSheet("Dashboard").Activate
    Exit Sub
Failure:
    MsgBox "Dashboard not found. Run InitializeInvoiceApp first.", vbExclamation
End Sub

Public Sub RefreshCatalogKpis()
    Dim dash As Worksheet
    On Error GoTo Failure
    Set dash = RequireSheet("Dashboard")
    dash.Range("B6").Value2 = CountActive(RequireSheet(CUSTOMERS))
    dash.Range("B7").Value2 = CountActive(RequireSheet(PRODUCTS))
    dash.Range("B8").Value2 = Date
    dash.Range("B8").NumberFormat = "yyyy-mm-dd"
    Exit Sub
Failure:
    MsgBox "Unable to refresh KPIs: " & Err.Description, vbExclamation
End Sub

Private Function CountActive(ByVal ws As Worksheet) As Long
    Dim r As Long
    For r = 2 To NextRow(ws) - 1
        If StrComp(Trim$(CStr(ws.Cells(r, 5).Value2)), "Active", vbTextCompare) = 0 Then
            CountActive = CountActive + 1
        End If
    Next r
End Function

Private Function NextRow(ByVal ws As Worksheet) As Long
    NextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If NextRow < 2 Then NextRow = 2
End Function

Private Function NextSequence(ByVal ws As Worksheet, ByVal prefix As String) As Long
    Dim r As Long, value As String, n As Long
    NextSequence = 1
    For r = 2 To NextRow(ws) - 1
        value = CStr(ws.Cells(r, 1).Value2)
        If Left$(value, Len(prefix)) = prefix Then
            If IsNumeric(Mid$(value, Len(prefix) + 1)) Then
                n = CLng(Mid$(value, Len(prefix) + 1))
                If n >= NextSequence Then NextSequence = n + 1
            End If
        End If
    Next r
End Function

Private Function RequireSheet(ByVal sheetName As String) As Worksheet
    Set RequireSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Public Sub WriteAudit(ByVal actionName As String, ByVal recordId As String, ByVal details As String)
    Dim ws As Worksheet, r As Long
    On Error GoTo SafeExit
    Set ws = RequireSheet("Audit Log")
    r = NextRow(ws)
    ws.Cells(r, 1).Value2 = Now
    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(r, 2).Value2 = actionName
    ws.Cells(r, 3).Value2 = recordId
    ws.Cells(r, 4).Value2 = details
SafeExit:
End Sub
