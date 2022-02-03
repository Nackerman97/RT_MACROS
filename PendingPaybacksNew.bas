Attribute VB_Name = "PendingPaybacksNew"
Option Explicit

'MAIN DATABASE
Private Const CorrespondenceConnectionString As String = "CONNECTION INPUT NEEDED"
'TESTING DATABASE
'Private Const CorrespondenceConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\nacke08\OneDrive - Safeway, Inc\Documents\Projects\BILLING_CORRESPONDENCE_BACKUP_01_14_22.mdb;Persist Security Info=False;"
'COVER_SHEET FILEPATH
Private Const CoversheetFilepath As String = "K:\AA\SHARE\AuditTools\rtmacros\data\PaybackCoversheet.xlsx"
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////MARKS Pending Payback TOOL//////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 06/23/2021
'Purpose:
'NOTES: Look for prepared by David Lehn = Paybacks Approved and Denied. Lindsay needs to update the comments. THis will be an update function.
'Update Function, Pull Function, Email Function
Public Sub OpenPaybacksEmail()
    Dim EmailBody As String
    Dim Data As Scripting.Dictionary
    Dim Todays_Date As String
    Dim OpenPaybacks As Integer
    
    LogCode "OpenPaybacksEmail"            'Log Reference to Tool
    
    'Check to make sure the sheet is open
    If ActiveWorkbook.name <> "Open Correspondence.xlsx" Then
        MsgBox "Please Open the Open Correspondence Excel File before using this function"
        Exit Sub
    End If
    
    OpenPaybacks = CInt(Cells(Rows.Count, 1).End(xlUp).row) - 1
    Set Data = JsonParse(ReadTextFile(ENV.use("CONFIG_BC")))
    
    'CONNECT TO BILLING CORRESPONDENCE DATABSE and pull open paybacks
    EmailBody = FormatTemplateEmail("K:\AA\SHARE\AuditTools\rtmacros\data\EmailTemplates\OpenPaybacksEmail.txt")
    EmailBody = Replace(EmailBody, "(&PAYBACKS_PREAUDIT)", CStr(OpenPaybacks))
    EmailBody = Replace(EmailBody, "(&PAYBACKS_MANAGER)", CStr(OpenPaybacks))
    
    'Create EMAIL
    BasicOutlookEmailCC JSONArrayUsernameToEmailString(ENV.use("CONFIG_BC"), "EmailUsers"), JSONArrayUsernameToEmailString(ENV.use("CONFIG_BC"), "CCEmailUsers"), "Open PB's", EmailBody

End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/09/2021
'Purpose: The purpose of this function is to take a query and place the query into a active worksheet.
'Only works with invoice details Table
Public Sub PullOpenPaybacksFromBC()
    'Step 1:Check if no one is in DB
    If EnsureBillingCorrespondenceDBIsClosedMsg = False Then Exit Sub
    'Step 2: Check Worksheet name
    If ActiveWorkbook.name <> "Open Correspondence.xlsx" Then
        MsgBox "Wrong Worksheet is Activated"
        Exit Sub
    Else
        Cells.Clear
    End If
    'Step 3:
    PullBillingCorrespondenceIntoWorksheet ReadTextFile("K:\AA\SHARE\AuditTools\rtmacros\sql\BillingCorrespondence\PullOpenPaybacks.sql")
    'Step 4:
    FormatOpenCorrespondenceSheet
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////Nancy/Sherri/Sean Billing Pending Payback TOOL///////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/18/2021
'Purpose: Designed for Nancy's team to pull Invoice numbers for Billers. This module is called by the application and will pull a invoice for a Billers
Public Sub PullBillingInvoiceNumber(ByVal InvoiceNumber As String)
    Dim SQL As String
    SQL = ReadTextFile("K:\AA\SHARE\AuditTools\rtmacros\sql\BillingCorrespondence\PullBillerInvoice.sql")
    SQL = Replace(SQL, "(&INVOICE_#)", InvoiceNumber)
    
    CommandsBillingInvoices SQL
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/18/2021
'Purpose: Designed for Nancy's team to pull Invoice numbers for Billers. This will accept a biller name and a divison string. The divison string is optional
Public Sub PullBillingInvoicesPreparedBy(ByVal Biller As String, ByVal Division As String, ByRef StatusType As String)
    Dim SQL As String
    
    'Replace and fill in template SQL strings
    If Division <> vbNullString Then
        SQL = ReadTextFile("K:\AA\SHARE\AuditTools\rtmacros\sql\BillingCorrespondence\PullPaybacksByPreparedBywDivision.sql")
        SQL = Replace(SQL, "(&DIVISION)", Division)
        SQL = Replace(SQL, "(&PREPAREDBY)", Biller)
        SQL = Replace(SQL, "(&SELECTION)", StatusType)
    Else
        SQL = ReadTextFile("K:\AA\SHARE\AuditTools\rtmacros\sql\BillingCorrespondence\PullPaybacksByPreparedBywDivision.sql")
        SQL = Replace(SQL, "(&PREPAREDBY)", Biller)
        SQL = Replace(SQL, "AND [INVOICE DETAIL].DivNo IN ((&DIVISION))", "")
        SQL = Replace(SQL, "(&SELECTION)", StatusType)
    End If
    
    CommandsBillingInvoices SQL
End Sub

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/18/2021
'Purpose: Designed for Nancy's team to pull Invoice numbers for Billers. This will pull all paybacks by division.
Public Sub PullBillingInvoicesByDivision(ByVal Division As String, ByRef StatusType As String)
    Dim SQL As String
    SQL = ReadTextFile("K:\AA\SHARE\AuditTools\rtmacros\sql\BillingCorrespondence\PullPaybacksByDivision.sql")
    SQL = Replace(SQL, "(&DIVISION)", Division)
    SQL = Replace(SQL, "(&SELECTION)", StatusType)
    
    CommandsBillingInvoices SQL
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/18/2021
'Purpose:
Public Sub CommandsBillingInvoices(ByRef SQL As String)
On Error GoTo err
    'Step 1:
    Sheets.Add
    'Step 2:
    PullBillingCorrespondenceIntoWorksheet SQL
    'Step 3:
    FormatBillingInvoiceSheet
    'Step 4:
    BillingPaybacks.CloseBillingPaybacks
err:
    Exit Sub
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/18/2021
'Purpose:This sub is used to build a coversheet for BillerPayback invoices. The user selects the row in Excel and then the coversheet is generated.
'place the coversheet values in the JSON
Public Sub CreatePaybackCoversheetV2()
On Error GoTo err
    Dim selectedRow As Long
    selectedRow = Selection.row
    
    LogCode "CreatePaybackCoversheetBilling"            'Log Reference to Tool
    
    'Step 1: check if directory exists and if it does not then create it
    If Dir("C:\Users\" & Environ("Username") & "\Documents\CoverSheets", vbDirectory) = "" Then
         MkDir "C:\Users\" & Environ("Username") & "\Documents\CoverSheets"
    End If
    
    'Step 2: check if selected row is valid
    If selectedRow <= 1 Or (Cells(selectedRow, FindHeading(ActiveSheet, "INVOICE #").column).Value) = vbNullString Then
        MsgBox "Please select a valid row"
        Exit Sub
    End If
    
    'Step 3: TODO ADD THESE FIELDS to the coversheet template JSON & Notify user of missing headers from file
    Application.ScreenUpdating = False
    Dim CoverSheetWorkbook As Excel.Workbook
    Dim BillingPaybackWorksheet As Worksheet
    Dim missingFields As String
    Set BillingPaybackWorksheet = ActiveSheet
    Set CoverSheetWorkbook = Workbooks.Open(CoversheetFilepath, True, True)
    missingFields = PopulatePaybackCoversheetFields(CoverSheetWorkbook, BillingPaybackWorksheet, selectedRow)

    'Step 4: Save as PDF (just using a temporary file so we don't have to upkeep a folder)
    With CoverSheetWorkbook.Sheets("Coversheet")
        .ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    End With
    
    'Step 5:Close workbook after it is saved And Reactive display updates
    CoverSheetWorkbook.Close (False)
    Application.ScreenUpdating = True
    
    'Step6:Display missing Fields from Coversheet.
    If missingFields <> vbNullString Then
        MsgBox "Raw Data is Missing the Filds Below" & vbNewLine & missingFields
    End If
    
    'Step 7:CurrentWorkbook.Activate
    MsgBox "Coversheets have been created", vbSystemModal
    Exit Sub
err:
    MsgBox "And Error Occured while Creating Coversheet" & err.Description
    Exit Sub
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/24/2021
'Purpose: The purpose of this function is to populate the fields of the coversheet using the values and location from the JSON file. This sub also checks to see if any of the headers are missing from the file.
Private Function PopulatePaybackCoversheetFields(CoverSheetWorkbook As Workbook, BillingPaybackWorksheet As Excel.Worksheet, selectedRow) As String
    Dim Data As Scripting.Dictionary
    Dim fieldSet, key As Variant
    Dim msgString As String
    Set Data = JsonParse(ReadTextFile(ENV.use("CONFIG_BC")))
    
    For Each fieldSet In Data("CoverSheetFields")
        For Each key In fieldSet.Keys()
            If key = "WHSE/DSD" Then
                'Special Case to convert 2 to NO, 1 to YES, and 0 to Null
                If BillingPaybackWorksheet.Cells(selectedRow, FindHeading(BillingPaybackWorksheet, CStr(key)).column).Value = "2" Then
                    CoverSheetWorkbook.Sheets("Coversheet").Range(CStr(fieldSet(key))) = "No"
                ElseIf BillingPaybackWorksheet.Cells(selectedRow, FindHeading(BillingPaybackWorksheet, CStr(key)).column).Value = "1" Then
                    CoverSheetWorkbook.Sheets("Coversheet").Range(CStr(fieldSet(key))) = "Yes"
                Else
                    CoverSheetWorkbook.Sheets("Coversheet").Range(CStr(fieldSet(key))) = ""
                End If
            ElseIf InStr(arrayFunctions.ArrayFromRow(1, BillingPaybackWorksheet).ToString, key) >= 1 Then
                If key = "ACCOUNT #" Then
                    CoverSheetWorkbook.Sheets("Coversheet").Range(CStr(fieldSet(key))) = BillingPaybackWorksheet.Cells(selectedRow, FindHeading(BillingPaybackWorksheet, CStr(key)).column).Value & "   " & BillingPaybackWorksheet.Cells(selectedRow, FindHeading(BillingPaybackWorksheet, "DEPARTMENT").column).Value
                Else
                    CoverSheetWorkbook.Sheets("Coversheet").Range(CStr(fieldSet(key))) = BillingPaybackWorksheet.Cells(selectedRow, FindHeading(BillingPaybackWorksheet, CStr(key)).column).Value
                End If
            Else
                msgString = msgString & "Missing " & CStr(key) & vbNewLine
            End If
        Next key
    Next fieldSet
    
    PopulatePaybackCoversheetFields = msgString
    
End Function
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////HELPER FUNCTIONS FOR BILLING PAYBACKS TOOLS///////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/19/2021
'Purpose: This function is to be used for the reports that are pulled on the billing end. This function formats the sheet to the liking of the billing teams.
'included is the ability to have drop downs listed for cells in the worksheet. Also the invoice column is locked and the header. This is to avoid someone
'accidently altering the column headers. The headers are bolded, the sheet is autosized
'TODO Find a way to make sure if a field exists before modifying it.
Public Sub FormatBillingInvoiceSheet()
    Dim numrows, row As Long
    numrows = CLng(Cells(Rows.Count, 1).End(xlUp).row)
    
    'Formatting of Worksheet
    Range("A1:BZ10000" & CStr(numrows)).EntireColumn.AutoFit  'autofit to cells
    Cells(1, 1).EntireRow.Font.Bold = True 'bold the row
    FindHeading(ActiveSheet, "INVOICE DATE").EntireColumn.NumberFormat = "mm/dd/yyyy" 'These need to be considered somehow in JSON
    FindHeading(ActiveSheet, "DATE OF VENDOR INQUIRY").EntireColumn.NumberFormat = "mm/dd/yyyy"
    FindHeading(ActiveSheet, "PAYBACK DATE").EntireColumn.NumberFormat = "mm/dd/yyyy"
    Range("A1:BZ10000").sort Key1:=Range(FindHeading(ActiveSheet, "STATUS").Address), Order1:=xlAscending, Header:=xlYes
    Range(FindHeading(ActiveSheet, "VENDOR REASON").Address).NoteText "REBILL AMOUNT" ' add notes to column
    Range(FindHeading(ActiveSheet, "WHSE/DSD").Address).NoteText "SHARED  1=Yes & 2 =No" ' add notes to column
    Range(FindHeading(ActiveSheet, "DSDBT").Address).NoteText "DSC - Direct Store Credit" ' add notes to column
    
    'Place filter on data
    Range("A1:" & ColumnNumber2Letter(GetColumnNumFromValue("CATEGORY MANAGER")) & "1").AutoFilter
   
    'Adding Dropdown fields to Excel cells' this is not formatting!!!!!!!
    For row = 2 To numrows
        ExcelDropDownFieldJSON row, FindHeading(ActiveSheet, "VENDOR REQUEST FOR PAYBACK").column, "CONFIG_BC", "VENDOR_REQUEST"
        ExcelDropDownFieldJSON row, FindHeading(ActiveSheet, "INT_REASON_CD").column, "CONFIG_BC", "INT_REASON_CD"
        ExcelDropDownFieldJSON row, FindHeading(ActiveSheet, "DivNo").column, "CONFIG_BC", "DivNo"
        ExcelDropDownFieldJSON row, FindHeading(ActiveSheet, "STATUS").column, "CONFIG_BC", "STATUS"
        ExcelDropDownFieldJSON row, FindHeading(ActiveSheet, "WHSE/DSD").column, "CONFIG_BC", "WHSE/DSD"
        ExcelDropDownFieldJSON row, FindHeading(ActiveSheet, "DSDBT").column, "CONFIG_BC", "DSDBT"
        ExcelDropDownFieldJSON row, FindHeading(ActiveSheet, "DEPARTMENT").column, "CONFIG_BC", "DEPARTMENT"
        'add a fomula to the invoice column with a hyperlink to the invoice in the R drive
        'ADD A FUNTION THAT WILL GO FOR APPROVED OR DENIED THEN POPUILATE A PAY DATE - USE AND IF THAN EQUATION and less than the invoice date
        Dim hyperlink, OFFER, DIV, invoice As String
        DIV = Replace(Cells(row, GetColumnNumFromValue("DivNo")).Value, " ", "")
        invoice = Replace(Cells(row, GetColumnNumFromValue("INVOICE #")).Value, " ", "")
        
        'CABS offers do not have hyperlinks 12/30/21
        If Left(Cells(row, GetColumnNumFromValue("INVOICE #")).Value, 3) <> "ALW" Then
            OFFER = Left(Cells(row, GetColumnNumFromValue("Offer_Num")).Value, 7)
            hyperlink = "R:\" & Left(OFFER, 4) & "xxx\" & OFFER & "\" & OFFER & "_" & DIV & "_" & invoice & "_PB.pdf"
            Cells(row, GetColumnNumFromValue("INVOICE #")).Value = "=HYPERLINK(" & Chr(34) & hyperlink & Chr(34) & "," & invoice & ")"
        End If
        
        'Add the Payback Date if the results do not equal PENDING
        Cells(row, GetColumnNumFromValue("PAYBACK DATE")).Value = "=IF(" & Right(GetColumnRngFromValue("VENDOR REQUEST FOR PAYBACK"), 1) & CStr(row) & "=" & Chr(34) & "PENDING" & Chr(34) & "," & Chr(34) & Chr(34) & ", TODAY())"
        
    Next row
    
    'Locking Sheet for Unintentional Editing
    ActiveSheet.Range("A1:BZ1").Locked = True
    ActiveSheet.Range("A2:BZ" & CStr(numrows)).Locked = False
    ActiveSheet.Range(FindHeading(ActiveSheet, "INVOICE #").Address & ":" & Left(FindHeading(ActiveSheet, "INVOICE #").Address, 3) & CStr(numrows)).Locked = True
    ActiveSheet.Protect "Billing123", AllowFormattingCells:=True, AllowFiltering:=True 'Passcode to unlock sheet (non critical)
    
End Sub
'This is a helper function that can be used for open correspondence or any query that utilzied the Payback columns. This can be added too for other fields in the Invoice details table
Private Sub FormatOpenCorrespondenceSheet()
    'Find Invoice Number Location
    Dim findCol, numrows, row As Long
    findCol = FindHeading(ActiveSheet, "VENDOR REQUEST FOR PAYBACK").column
    numrows = CLng(Cells(Rows.Count, 1).End(xlUp).row)
    
    'Formatting of Worksheet
    Range("A1:BZ" & CStr(numrows)).EntireColumn.AutoFit 'autofit to cells
    Cells(1, 1).EntireRow.Font.Bold = True 'bold the row
    
    For row = 2 To numrows
        'DropDown SELECTION For the VENDOR REQUEST FOR PAYBACK COLUMN
        ExcelDropDownFieldJSON row, FindHeading(ActiveSheet, "VENDOR REQUEST FOR PAYBACK").column, "CONFIG_BC", "VENDOR_REQUEST"
        'Cells(row, GetColumnNumFromValue("PAYBACK DATE")).value = "=DATE(YEAR(NOW()),MONTH(NOW()), DAY(NOW()))"
        Range(GetColumnRngFromValue("PAYBACK DATE")).NumberFormat = "dd/mm/yyyy"
    Next row
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/19/2021
'Purpose: This is a helper function that is used to place a dropdown field into a cell that is defined by the row
'col by the user. The=is function requires that a JSONfile be sent to it that has a reference key in the evnironment variables. It also
'requires the field name fro the JSON string. The field name must be a string and not an array.
Public Sub ExcelDropDownFieldJSON(ByVal row As Long, ByVal col As Long, ByRef JSONFilename As String, ByRef JSONFieldname As String)
    Dim Data As Scripting.Dictionary
    Set Data = JsonParse(ReadTextFile(ENV.use(JSONFilename)))
    
    With Cells(row, col).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Data(JSONFieldname)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
    End With
    
End Sub

'This is helper function that prompts a message box for users so that they are not using the billing correspondenc databse if a user is already in the DB.
Private Function EnsureBillingCorrespondenceDBIsClosedMsg() As Boolean
    Dim ans As String
    ans = MsgBox("Have you check to ensure that no one is in the Billing Correspondence DB. OK = YES, CANCEL = NO :(", vbOKCancel)
    If ans = vbOK Then
        EnsureBillingCorrespondenceDBIsClosedMsg = True
    Else
        EnsureBillingCorrespondenceDBIsClosedMsg = False
    End If
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/09/2021
'Purpose: The purpose of this function is to take a query and place the query into a active worksheet.
'Only works with invoice details Table in the billing correspondence database. An invopice number must be included with the query to be
'placed in the table
Public Sub PullBillingCorrespondenceIntoWorksheet(ByRef SQL As String)
On Error GoTo err
    'CONNECTING TO ACCESS DB
    Dim objRec As ADODB.Recordset
    Dim objConn As Object

    Set objRec = CreateObject("ADODB.Recordset")
    Set objConn = CreateObject("ADODB.Connection")
    objConn.connectionString = CorrespondenceConnectionString
    objConn.Open

    Set objRec = objConn.Execute(SQL)
    
    'Check if RecordSet is empty
    If objRec.EOF = True Then
        GoTo records
    End If
    
    'pull records set of data with headers into the worksheet
    RecordSetFunctions.RecordsetToRange objRec, ActiveSheet.Range("A1")
    
    objConn.Close
    MsgBox "SUCCESS! BILLING CORRESPONDENCE DATA HAS BEEN PULLED"
    Exit Sub
    
err:
    objConn.Close
    MsgBox "And Error Occured while updating the Billing Correspondence Database " & err.Description
    Exit Sub
    
records:
    objConn.Close
    MsgBox "No Records Found"
    Exit Sub
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/24/2021
'Purpose: The purpose of this function is to take the data from the Billing Correspondence worksheet that is being used and update the values that exists with
'The database. This function is universal and does not need a query to operate. All it needs is the worksheet to contain the proper headers for the billing correspondence database
'and the invoice number.
Public Sub UpdateBillingCorrespondenceFromWorksheet()
On Error GoTo err
    Dim numrows, numCols, col, row, invoiceCol As Long
    Dim objConn As Object
    Dim SQL As String
    numCols = CLng(Cells(1, Columns.Count).End(xlToLeft).column)
    invoiceCol = FindHeading(ActiveSheet, "INVOICE #").column
    numrows = CLng(Cells(Rows.Count, invoiceCol).End(xlUp).row)
    
    'Check if Correspondence DB is CLOSED
    If EnsureBillingCorrespondenceDBIsClosedMsg = False Then Exit Sub
    
    Set objConn = CreateObject("ADODB.Connection")
    objConn.connectionString = CorrespondenceConnectionString
    objConn.Open
    
    For row = 2 To numrows
        objConn.Execute BuildUpdateQueryForBillingCorrespondence(numCols, invoiceCol, col, row)
    Next row
    
    objConn.Close
    MsgBox "SUCCESS! BILLING CORRESPONDENCE DATABASE HAS BEEN UPDATED"
    Exit Sub
    
err:
    objConn.Close
    MsgBox "And Error Occured while updating the Billing Correspondence Database - " & err.Description '
    Exit Sub
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/24/2021
'Purpose: This is a helper function for UpdateBillingCorrespondenceFromWorksheet. This function produces the SQL that is needed for the main function. The idea is this makes
'the funciton more readable. Could also move the other loop into this Function
Private Function BuildUpdateQueryForBillingCorrespondence(ByVal numCols As Long, ByRef invoiceCol As Long, ByVal col As Long, ByVal row As Long) As String
    Dim SQL As String
    SQL = "UPDATE [INVOICE DETAIL] SET "
    
    For col = 1 To numCols
        If (col <> invoiceCol) And (Cells(row, col).Value <> vbNullString) Then
            If UCase(Cells(row, col).Value) = "NULL" Then
                SQL = SQL & "[INVOICE DETAIL]." & "[" & Cells(1, col).Value & "] = NULL, "
            Else
                SQL = SQL & "[INVOICE DETAIL]." & "[" & Cells(1, col).Value & "] = '" & RemoveSpecialChars(Cells(row, col).Value) & "', "
            End If
        End If
    Next col
    SQL = Left(SQL, Len(SQL) - 2)
    SQL = SQL & " WHERE [INVOICE DETAIL].[INVOICE #] = '" & CStr(Cells(row, invoiceCol).Value) & "';"
    'Debug.Print SQL
    BuildUpdateQueryForBillingCorrespondence = SQL
End Function

