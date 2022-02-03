Attribute VB_Name = "Tribbler"
Option Explicit
Private mz As New materialize

'Created 08/25/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
Sub TribbleProcess()
    TribbleForm.show                                                                        'Open the Tribble Form
End Sub


'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 09/08/2020
'Purpose:This is called after the successful creation of a Tribble
'References:FormatTemplateEmail,Replace,JsonParse
Public Sub ProduceTribbleEmail(DIV As String, WIMS_Vndr_Num As String, issue As String, VENDOR_NAME As String, Tribble_ID As String, OfferNumber As String, ASM As String, FIRSTCIC As String)
On Error GoTo catch
    Dim EmailBody As String
    Dim subjectHeader As String
    Dim Data As Scripting.Dictionary
    Dim item As Variant

    Set Data = JsonParse(ReadTextFile(ENV.use("CONFIG_TRIBBLE")))

    MsgBox "Emails are being generated"

    EmailBody = FormatTemplateEmail(ENV.use("EMAILTRIBBLERFOLDERPATH") & "\" & issue & ".txt")
    EmailBody = Replace(EmailBody, "(&DIV)", DIV)
    EmailBody = Replace(EmailBody, "(&WIMS_VNDR_NUM)", WIMS_Vndr_Num)
    EmailBody = Replace(EmailBody, "(&ISSUE)", issue)
    EmailBody = Replace(EmailBody, "(&VENDOR_NAME)", VENDOR_NAME)
    EmailBody = Replace(EmailBody, "(&TRIBBLE_ID)", Tribble_ID)
    EmailBody = Replace(EmailBody, "(&OFFER_NUM)", OfferNumber)
    EmailBody = Replace(EmailBody, "(&ASM)", ASM)
    EmailBody = Replace(EmailBody, "(&CIC)", FIRSTCIC)
    EmailBody = Replace(EmailBody, "<hglt>", "<span style='background:yellow;mso-highlight:yellow'>")                   'For highlighting text
    
    'if of a specific issue type add more to the body of the email
    If issue = "OINR" Then
        EmailBody = OINREmailBody(EmailBody, Tribble_ID)
    End If
    
    'If offer number is 0 then replace with CIC value
    If OfferNumber = "00000000" Or OfferNumber = "0" Then
        subjectHeader = issue & " - CIC " & FIRSTCIC & " - " & VENDOR_NAME & "- D" & DIV & " - TribbleID:" & Tribble_ID
    Else
        subjectHeader = issue & " - Offer " & OfferNumber & " - " & VENDOR_NAME & " D" & DIV & " - TribbleID:" & Tribble_ID
    End If
    
    'Condition for Bouncer Emails
    If issue = "OINR" Then
        BasicOutlookEmail "", subjectHeader, EmailBody
    ElseIf Contains(Data("bouncers"), Environ("Username")) = True Then
        BasicOutlookEmail "marie.vaughan@albertsons.com", subjectHeader, EmailBody
    Else
        BasicOutlookEmail "Preaudit.Request@albertsons.com", subjectHeader, EmailBody
    End If
    
    'Check to see if Tribble DB is ready for Archiving
    CheckIfTribbleArchiveIsReady
    
'Error Handling
catch:
    If err.Number <> 0 Then
        CatchErrorController "Tribbler.ProduceTribbleEmail"
        Exit Sub
    End If
End Sub


'Calls an sql query from the DB2 server to pull NOPA information.
'This inforamtion is then populated in the New Tribble Form.
Public Sub AutoFillerTesting(OfferNum As Long)
    On Error GoTo catch
    Dim Data As Variant
    Dim SQL As String
    'Dim DIV As String
    
    'DIV = InputBox("Please enter a DIVSION#: ")
    'VALIDATE_TRIBBLE_DATA
    
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\Get_Nopa_Info_V2.sql")
    SQL = Replace(SQL, "(&Offer_Num)", OfferNum)
    
    Data = DB2.RunQuery(SQL)
    
    Dim Search As String
    Dim item As Variant
    
    'ADD CICS TO LIST BOX BEFORE POPULATING OT ENSURE NO DUPLICATES
    
    For Each item In Data
        Select Case True
            Case item = "RTL_ITM_NBR":      Search = "CICS"
            Case item = "PRFRM_START_DT":   Search = "StartDate"
            Case item = "PRFRM_END_DT":     Search = "EndDate"
            Case item = "DIV":              Search = "Div"
            Case item = "ASM":              Search = "ASM"
            Case item = "VEND_NUM":         Search = "WIMS_VEND_NUM"
            Case item = "VENDOR_NAME":      Search = "VendNM"
            Case Search = "CICS" And CheckIfCICValueExistsInListBox(item) = False:            TribbleFormNew.CicsListBox.AddItem item
            Case Search = "StartDate":      TribbleFormNew.StartTextBox.Value = item
            Case Search = "EndDate":        TribbleFormNew.EndTextBox.Value = item
            Case Search = "VendNM":         VendorNameClearExtraSpaces (item)
            Case Search = "WIMS_VEND_NUM":  TribbleFormNew.WimsTextBox.Value = item
            Case Search = "ASM":            TribbleFormNew.ASMTextBox.Value = item
            Case Search = "Div":            TribbleFormNew.FillDivisionBoxes (CStr(item))
        End Select
    Next item
    
    Exit Sub
'Error Handling
catch:
    If err.Number <> 0 Then
        CatchErrorController "Tribbler.AutoFillerTesting"
        Exit Sub
    End If
    
End Sub

'Used when populating a new Tribble in TribbleFormNew
'New Tribbles have their CICS and divison compared.
'This turns the list of CICS in the list box into a String (1111,2222,3333).
Public Function ProduceTribbleCicsString() As String
    On Error GoTo catch
    Dim i As Integer
    Dim CicsString As String
    CicsString = "("
    For i = 0 To TribbleFormNew.CicsListBox.ListCount - 1
        If i = (TribbleFormNew.CicsListBox.ListCount - 1) Then
            CicsString = CicsString & TribbleFormNew.CicsListBox.List(i)
        Else
            CicsString = CicsString & TribbleFormNew.CicsListBox.List(i) & ", "
        End If
    Next i
    CicsString = CicsString & ")"
    
    ProduceTribbleCicsString = CicsString
    
    Exit Function
'Error Handling
catch:
    If err.Number <> 0 Then
        CatchErrorController "Tribbler.ProduceTribbleCicsString"
        Exit Function
    End If
    
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 01/05/2021
'Purpose: To retrieve the last TribbleID that exists in the DB
'References:ReadTextFile,IsRecordsetEmpty QuerySnowFlake
Public Function GetLastTribbleID() As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\Get_Last_Tribbler.sql") 'find the last tribbler
    Set rs = QuerySnowFlake(SQL)
    
    If IsRecordsetEmpty(rs) Then
        Exit Function
    End If
    
    Do While Not rs.EOF
        GetLastTribbleID = rs.Fields("TRIBBLE_ID").Value
    rs.MoveNext
    Loop
    
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 09/25/2020
'Purpose:'Used to divide up workload. JSON file lists the Tribblers
'References:JsonParse,QuerySnowFlake,ReadTextFile,ArrayIndexOf
Public Function TribblerRoundRobin() As String
On Error GoTo catch
    Dim Data As Scripting.Dictionary
    Dim item As Variant
    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim TRIBBLER As String

    Set Data = JsonParse(ReadTextFile(ENV.use("CONFIG_TRIBBLE")))

    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\Get_Last_Tribbler.sql") 'find the last tribbler
    Set rs = QuerySnowFlake(SQL)

    If IsRecordsetEmpty(rs) Then
        Exit Function
    End If

    Do While Not rs.EOF
        If rs.Fields("TRIBBLER").Value = "NULL" Then
            TRIBBLER = "Mark"
        Else
            TRIBBLER = rs.Fields("TRIBBLER").Value
        End If
        
    rs.MoveNext
    Loop
    
    'index of last tribbler in json file
    Dim index As Integer
    index = ArrayIndexOf(Data("tribbler"), TRIBBLER)

    'pick next tribbler in json file
    If index = Length(Data("tribbler")) - 1 Then
        TRIBBLER = Data("tribbler")(0)
    Else
        TRIBBLER = Data("tribbler")(index + 1)
    End If

    TribblerRoundRobin = TRIBBLER

    Exit Function
'Error Handling
catch:
    If err.Number <> 0 Then
        CatchErrorController "Tribbler.TribblerRoundRobin"
        Exit Function
    End If

End Function

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 09/29/2020
'Purpose:'Used as the controller for error statements
'References:DisplayErrorMessage
Public Function CatchErrorController(ErrorString As String)
    Dim ErrorMessage As String
    ErrorMessage = err.Description

    Application.ScreenUpdating = True
    Console.error ErrorMessage, ErrorString
    
    With New TribbleErrorForm
        .ErrorMessage = ErrorString + " " + ErrorMessage
        .show
    End With
    
End Function

'TribbleFormNew.VendorNameTextBox.value = item
'credits StackOverflow User :)
'Created: 10/05/2020
'Purpose: Eliminate proceeding and proceeeding spaces
Public Sub VendorNameClearExtraSpaces(VendorName As String)
    Dim RE As RegExp
    Set RE = New RegExp
    With RE
        .Global = True
        .MultiLine = True
        .pattern = "^\s*(\S.*\S)\s*"
        TribbleFormNew.VendorNameTextBox.Value = .Replace(VendorName, "$1")
    End With
End Sub

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/31/2020
'Purpose: The purpose of this function is to delete the old Cics in the database an populate them with new Cic values
'References:Access connection for deleting CICS
Public Sub DeleteOldCics(Tribble_ID As String)
    Dim item As Variant
    Dim SQL As String
    Dim err As Integer

    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\New_Delete_CICS.sql")
    SQL = Replace(SQL, "(&Tribble_ID)", Tribble_ID)
    
    err = InsertDataIntoSnowflake(SQL)
    If err <> 0 Then
        MsgBox "And Error occured Deleting Old CICS. Please contact BA"
    End If
End Sub

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 09/17/2020
'Purpose:This is called after pressing the button ton the main screen of RT amcros.
'References:GetRawTribbleDataReport
Public Sub TribbleReporting()
    On Error GoTo catch
    
    Dim wb As Workbook
    Set wb = Workbooks.Open("K:\AA\SHARE\AuditTools\rtmacros\data\Reports\Tribble_Report.xlsm")
    
    Exit Sub
'Error Handling
catch:
    If err.Number <> 0 Then
        CatchErrorController "Tribbler.TribbleReporting"
        Exit Sub
    End If
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 11/06/2020
'Purpose:Generates a query based off of the user selection Form
Public Function GenerateTribbleQuery(TRIBBLER As String, Status As String) As String
    Dim SQL As String
   
    If TRIBBLER <> vbNullString And Status <> vbNullString Then
        SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\All_Status_Tribbler.sql")
        SQL = Replace(SQL, "TRIBBLER ='(&Tribbler)'", "PREAUDITOR='" & TRIBBLER & "'")
        SQL = Replace(SQL, "(&Status)", Status)
    ElseIf TRIBBLER = vbNullString And Status <> vbNullString Then
        SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\All_Status.sql")
        SQL = Replace(SQL, "(&Status)", Status)
        
    ElseIf TRIBBLER <> vbNullString And Status = vbNullString Then
        SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\All_Tribbler.sql")
        SQL = Replace(SQL, "TRIBBLER ='(&Tribbler)'", "PREAUDITOR='" & TRIBBLER & "'")
    Else
        SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\All_Types.sql")
    End If
    
    GenerateTribbleQuery = SQL
    
End Function

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 11/06/2020
'Purpose:Generates a query of all the Open Inquiry Tribbles available
Public Sub GenerateTribbleUpdateQuery()
    Dim AnswerYes As String
    Dim rs As ADODB.Recordset
    Dim subject As Variant
    Dim SQL As String
    
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\All_OpenInquiry.sql")
    Set rs = QuerySnowFlake(SQL)
    
    If IsRecordsetEmpty(rs) Then
        MsgBox "There are currently no Tribbles listed as 'Open Inquiry' or 'Email'"
        Exit Sub
    End If
    
     AnswerYes = MsgBox("Are you sure you want to Update the DB and search for Emails?", vbQuestion + vbYesNo, "UpdateDB")
    
    If AnswerYes = vbYes Then
        Do While Not rs.EOF
            SearchEntirePreAuditMailbox rs
        rs.MoveNext
        Loop
        MsgBox "The database has been updated"
    End If
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 02/10/2021
'Purpose:This function changes the automatically assigned tribbler to another user defined tribbler
'Referencecs: regex , JsonParse
'input: TribbleID
'Output: None
Public Sub ChangeTribbler(ByVal TribbleID As String)
    
    Dim Data As Scripting.Dictionary
    Set Data = JsonParse(ReadTextFile(ENV.use("CONFIG_TRIBBLE")))
    
    Dim User As Variant
    Dim Count As Integer
    Dim TransferMessage As String
    Dim JSONTribblers As Collection
    Set JSONTribblers = New Collection
    
    Count = 1
    'Places all of the Tribblers into a collection and Tribble reassign message
    For Each User In Data("tribbler")
        TransferMessage = TransferMessage & vbCr & "[" & Count & "]: " & User
        JSONTribblers.Add User
        Count = Count + 1
    Next User
    
    'Add BOUNCER LEAD TO TRIBBLER Message
    TransferMessage = TransferMessage & vbCr & "[" & Count & "]: " & Data("bouncer_lead")
    JSONTribblers.Add Data("bouncer_lead")
    
    
    'If Contains(Data("tribbler"), "mcart49") = True Then
    Dim TransferUID As String

    'input for UserID
    TransferUID = InputBox("Who do you want to transfer TribbleID: " & TribbleID & " too?" & vbCr & TransferMessage)
    If regularExpressionExists(TransferUID, "[1-" & JSONTribblers.Count & "]") = False Then
        GoTo process
    End If
    
    'Run a query to change the Tribbler for the tribble
    Dim SQL As String
    Dim err As Integer
    SQL = "UPDATE DW_PRD.TEMP_TABLES.PA_TRIBBLES SET TRIBBLER='(&Tribbler)' WHERE TRIBBLE_ID =(&TribbleId);"
    SQL = Replace(SQL, "(&Tribbler)", JSONTribblers.item(CInt(TransferUID)))
    SQL = Replace(SQL, "(&TribbleId)", TribbleID)
    err = InsertDataIntoSnowflake(SQL)
    
    'Check that change was completed
    If err = 0 Then
        MsgBox "Tribble: " & TribbleID & " has been assigned too: " & JSONTribblers.item(CInt(TransferUID))
    Else
        GoTo catch
    End If
                        
    'Else
    '    MsgBox "You do not have access to this function"
    'End If
    Exit Sub
    
process:
    MsgBox "The number you entered is invalid"
    Exit Sub
catch:
    MsgBox "An Error Occured Updating this Tribble. Please contact BA"
    Exit Sub
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 02/10/2021
'Purpose:Check if a Tribble Exists
'[True] either an error has occured or a Tribble exists
'[False] a tribble does not exists
'This function is not connected to anything yet!!!!!!!!!!!
Public Function CheckIfTribbleExists(ByVal TribbleID As String) As Boolean
    Dim SQL As String
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\New_Tribble_Check_Exists_DIV_CICS.sql")
    SQL = Replace(SQL, "(&DIVISION)", DIV)
    
    Dim rs As ADODB.Recordset
    Set rs = QuerySnowFlake(SQL)
    
    If IsRecordsetEmpty(rs) = True Then
        CheckIfTribbleExists = False
    End If

End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'updated: 07/19/2021
'Purpose:The purpose of this function is to open a .txt file located at C:\rtmacros.
'This will allow for an easier input of variables into the CIC box by the user.
'IF the file is not located here then a file is made on the users computer.
'The sql query checks to see if all of the CICS are valid in a larges search
'References: CICStringToCollection , FilterCICResults, AddValidCICsToForm, QuerySnowFlake, ReadTextFile
Public Sub CICUpload(ByRef DIV As String)
    'Check to see if FILE exists
    Dim fso As Object
    Dim ts As Object

    Dim CICList As String
    
    'GET UPCLIST
    If Not FileExists("C:\rtmacros\CICS.txt") Then
        MsgBox "There is no CICS.txt file on your desktop", vbInformation
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.OpenTextFile("C:\rtmacros\CICS.txt", 2, True)
        Exit Sub
    End If

    CICList = ReadTextFile("C:\rtmacros\CICS.txt")

    If Trim(CICList) = "" Then
        GoTo msgCICS
    Else
        Dim SQL As String
        SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\CICS_IMPORT_FILE.sql")
        SQL = Replace(SQL, "(&CICS)", CICList)
        SQL = Replace(SQL, "(&DIV)", DIV)
        
        Dim rs As ADODB.Recordset
        Set rs = QuerySnowFlake(SQL)

        If IsRecordsetEmpty(rs) = True Then
            AddValidCICsToForm CICStringToCollection(CICList)
        Else
            DupCICResultMessage rs
            AddValidCICsToForm CICStringToCollection(CICList)
        End If
        
        rs.Close
    End If
    Exit Sub
    
msgCICS:
    MsgBox "CICS.Txt on your desktop is empty", vbInformation
    Exit Sub
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 02/12/2021
'Purpose:
Private Sub DupCICResultMessage(ByVal rs As ADODB.Recordset)
    Dim messageString As String
    messageString = "TribbleID" & vbTab & vbTab & "CICS" & vbCr & vbCr

    Do While Not rs.EOF
        messageString = messageString & rs.Fields("Tribble_ID").Value & vbTab & vbTab & rs.Fields("CICS").Value & vbCr
        rs.MoveNext
    Loop
    
    MsgBox (messageString)
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Updated: 04/14/2021
'Purpose:
Private Function CICStringToCollection(ByVal CICList As String) As Collection
    Dim tempString() As String
    Set CICStringToCollection = New Collection
    tempString = split(CICList, vbNewLine)
    
    Dim temp As Variant
    For Each temp In tempString
        temp = Replace(temp, ",", "")
        CICStringToCollection.Add (Left(temp, Length(temp)))
    Next temp
    
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 02/12/2021
'Purpose:
Private Sub AddValidCICsToForm(ByVal CollectionCICS As Collection)
    Dim CIC As Variant
    For Each CIC In CollectionCICS
        TribbleFormNew.CicsListBox.AddItem CIC
    Next CIC
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 03/24/21
'Purpose: Possible reduction in code for populating a text file string with values
'Current issue is the same of the sql files do not contain the same replace variables
Public Function TribblerSQLStringReplace(ByRef SQL As String, Optional Tribble_ID As String, Optional TodaysDate As String, Optional DIV As String, Optional WIMSVendNbr As String, Optional ASM As String, _
Optional FIRSTCIC As String, Optional PREAUDITOR As String, Optional VendorName As String, Optional OfferNum As String, Optional TimeFrameStart As String, Optional TimeFrameEnd As String, _
Optional issue As String, Optional Status As String, Optional TRIBBLER As String, Optional TribbleNotes As String, Optional LastContactMade As String, Optional NumContact As String) As String

    SQL = Replace(SQL, "(&TRIBBLE_ID)", Tribble_ID)
    SQL = Replace(SQL, "(&Todays_Date)", TodaysDate)
    SQL = Replace(SQL, "(&DIV)", DIV)
    SQL = Replace(SQL, "(&WIMS_Vndr_Num)", WIMSVendNbr)
    SQL = Replace(SQL, "(&ASM)", ASM)
    SQL = Replace(SQL, "(&FirstCIC)", FIRSTCIC)
    SQL = Replace(SQL, "(&PreAuditor)", UCase(PREAUDITOR))
    SQL = Replace(SQL, "(&Vendor_Name)", VendorName)
    SQL = Replace(SQL, "(&Offer_Num)", OfferNum)
    SQL = Replace(SQL, "(&TimeFrame_Start)", TimeFrameStart)
    SQL = Replace(SQL, "(&TimeFrame_End)", TimeFrameEnd)
    SQL = Replace(SQL, "(&issue)", issue)
    SQL = Replace(SQL, "(&Status)", Status)
    SQL = Replace(SQL, "(&Tribbler)", UCase(TRIBBLER))
    SQL = Replace(SQL, "(&Tribble_Notes)", TribbleNotes)
    SQL = Replace(SQL, "(&Last_Contact_Made)", LastContactMade)
    SQL = Replace(SQL, "(&Num_Contact)", NumContact)
    
    TribblerSQLStringReplace = SQL
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 03/24/21
'Purpose: Possible reduction in code for populating same message box for two function in the
'TribblerFormNew module. This message box displays the tribbles that already exists. Input is a
'recordset and the output is a VBA cancel or VBA ok statement. The messages box is displayed in this function as well
Public Function TribblerTribbleExistsMsgBox(ByRef rs As Recordset) As String

    Dim MsgBoxString As String
    Dim row, Count As Long
    MsgBoxString = "Conflicting Tribbles. Do you Wish to Add the Tribble Anyway?" & vbCr & vbCr & _
       "[Cancel] to Exit" & vbCr & "[OK] to Add" & vbCr & vbCr
    
    Count = 1
    Do While Not rs.EOF
        If Count <= 3 Then
            MsgBoxString = MsgBoxString & "Tribble ID: " & vbTab & rs.Fields("TRIBBLE_ID").Value & vbCr & _
            "WIMS_Vndr_Num: " & vbTab & rs.Fields("WIMS_VNDR_NUM").Value & vbCr & _
            "Issue: " & vbTab & rs.Fields("ISSUE").Value & vbCr & _
            "Status: " & vbTab & rs.Fields("STATUS").Value & vbCr & _
            "Division: " & vbTab & rs.Fields("DIVISION").Value & vbCr & _
            "FirstCIC: " & vbTab & rs.Fields("FIRSTCIC").Value & vbCr & _
            "Vendor_Name: " & vbTab & rs.Fields("VENDOR_NAME").Value & vbCr & _
            "ASM: " & vbTab & rs.Fields("ASM").Value & vbCr & _
            "Timeframe_Start: " & vbTab & rs.Fields("TIMEFRAME_START").Value & vbCr & _
            "Timeframe_End: " & vbTab & rs.Fields("TIMEFRAME_END").Value & vbCr & _
            "--------------------------------------------------" & vbCr
            Count = Count + 1
            rs.MoveNext
        Else
            Exit Do
        End If
    Loop
    
    TribblerTribbleExistsMsgBox = MsgBox(MsgBoxString, vbQuestion + vbOKCancel, "User Repsonse")

End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 03/25/2021
'Purpose:The purpose of the function is to delete a tribble and its corresponding CICS from the snowflake DB
'based off of the associated TribbleID. This function is activated by the ListBox1_KeyDown function
Public Sub DeleteTribble(ByVal TribbleID As String)
    Dim ans As String
    Dim SQL As String
    ans = MsgBox("WARNING: You are about to delete TribbleID: " & TribbleID, vbQuestion + vbYesNo, "Delete Tribble")
    If ans = vbYes Then
        Dialog.show "Deleting Tribble...", "This might take a few seconds"
        SQL = "DELETE FROM DW_PRD.TEMP_TABLES.PA_TRIBBLES WHERE TRIBBLE_ID = '" & TribbleID & "'"
        InsertDataIntoSnowflake SQL
        SQL = "DELETE FROM DW_PRD.TEMP_TABLES.PA_TRIBBLES_CICS WHERE TRIBBLE_ID = '" & TribbleID & "'"
        InsertDataIntoSnowflake SQL
        MsgBox ("Tribble: " & TribbleID & " has been deleted")
        Dialog.Hide
    End If
End Sub
Public Function TribbleSearchString(SMONTH, EMONTH, CICS, TRIBBLER, issue, Status, WIMS, ASM, Vendor, TribbleID, OFFER, PREAUDITOR, DIV, Closed_Resolved) As String
    Dim SQL As String
    Dim SDATE As String
    Dim EDATE As String
    
    SDATE = CStr(YEAR(SMONTH)) & "-" & CStr(Month(SMONTH)) & "-" & CStr(Day(SMONTH)) & " 00:00:00"
    EDATE = CStr(YEAR(EMONTH)) & "-" & CStr(Month(EMONTH)) & "-" & CStr(Day(EMONTH)) & " 00:00:00"
    SQL = "SELECT * FROM DW_PRD.TEMP_TABLES.PA_TRIBBLES WHERE"
    
    'CIC
    If CICS <> vbNullString Then
        SQL = "SELECT DISTINCT T.TRIBBLE_ID, " & _
        "T.TODAYS_DATE, T.DIVISION, T.WIMS_VNDR_NUM, " & _
        "T.ASM, T.FIRSTCIC, T.PREAUDITOR, " & _
        "T.VENDOR_NAME, T.OFFER_NUM, T.TIMEFRAME_START, " & _
        "T.TIMEFRAME_END, T.ISSUE, T.STATUS, " & _
        "T.TRIBBLER, T.TRIBBLE_NOTES, T.LAST_CONTACT_MADE, " & _
        "t.NUM_CONTACT_MADE, t.NUM_EMAILS " & _
        "FROM DW_PRD.TEMP_TABLES.PA_TRIBBLES T " & _
        "LEFT JOIN TEMP_TABLES.PA_TRIBBLES_CICS C " & _
        "ON T.TRIBBLE_ID =c.TRIBBLE_ID " & _
        "WHERE c.CICS = " & CICS
    End If
    
    'Tribbler
    If TRIBBLER <> vbNullString Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND TRIBBLER = '" & TRIBBLER & "' "
        Else
            SQL = SQL & " TRIBBLER = '" & TRIBBLER & "' "
        End If
    End If
    
    'ISSUE
    If issue <> vbNullString Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND ISSUE = '" & issue & "' "
        Else
            SQL = SQL & " ISSUE = '" & issue & "' "
        End If
    End If
    
    'STATUS
    If Status <> vbNullString Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND STATUS = '" & Status & "' "
        Else
            SQL = SQL & " STATUS = '" & Status & "' "
        End If
    End If
    
    'WIMS_VNDR_NUM
    If WIMS <> vbNullString Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND WIMS_VNDR_NUM = '" & WIMS & "' "
        Else
            SQL = SQL & " WIMS_VNDR_NUM = '" & WIMS & "' "
        End If
    End If
    
    'ASM
    If ASM <> vbNullString Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND UPPER(ASM) LIKE '%" & UCase(ASM) & "%' "
        Else
            SQL = SQL & " UPPER(ASM)LIKE '%" & UCase(ASM) & "%' "
        End If
    End If
    
    'VENDOR_NAME
    If Vendor <> vbNullString Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND UPPER(VENDOR_NAME) LIKE '%" & UCase(Vendor) & "%'"
        Else
            SQL = SQL & " UPPER(VENDOR_NAME)LIKE '%" & UCase(Vendor) & "%'"
        End If
    End If
    
    'Tribble ID
    If TribbleID <> vbNullString Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND TRIBBLE_ID = " & TribbleID
        Else
            SQL = SQL & " TRIBBLE_ID = " & TribbleID
        End If
    End If
    
    'OFFER_NUM
    If OFFER <> vbNullString Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND OFFER_NUM = '" & OFFER & "' "
        Else
            SQL = SQL & " OFFER_NUM = '" & OFFER & "' "
        End If
    End If
    
    'PreAuditor
    If PREAUDITOR <> vbNullString Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND PREAUDITOR = '" & PREAUDITOR & "' "
        Else
            SQL = SQL & " PREAUDITOR = '" & PREAUDITOR & "' "
        End If
    End If
    
    'DIVISION
    If DIV <> vbNullString Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND (" & DIV & ") "
        Else
            SQL = SQL & " (" & DIV & ") "
        End If
    End If
    
    'Status with or without CLOSED/RESOLVE
    If Closed_Resolved = False Then
        If Len(SQL) >= 55 Then
            SQL = SQL & " AND STATUS != 'Closed' AND STATUS != 'Closed-Preaudit' AND STATUS != 'Resolved'"
        Else
            SQL = SQL & " STATUS != 'Closed' AND STATUS != 'Closed-Preaudit' AND STATUS != 'Resolved'"
        End If
    End If
    
    'SDATE AND EDATE
    If Len(SQL) >= 55 Then
        SQL = SQL & " AND (TODAYS_DATE BETWEEN '" & SDATE & "' AND '" & EDATE & "')"
    Else
        SQL = SQL & " (TODAYS_DATE BETWEEN '" & SDATE & "' AND '" & EDATE & "')"
    End If
    
    SQL = SQL & " ORDER BY LAST_CONTACT_MADE DESC"
    
    TribbleSearchString = SQL
    
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 04/16/2021
'PURPOSE: The purpose of this function is to serve as the header function for both AddCICSToEmail and AddOITableToEmail.
'Primariliy the return values from those function are placed into the trmplate text file for the email. This function also take in two
'user inputs for the PO number and the facility to generate the email. These values are also autoamtically saved into the notes section for
'the tribble
'REFERENCES: AddOITableToEmail, AddCICSToEmail, TribblerSQLStringReplace, InsertDataIntoSnowflake
Private Function OINREmailBody(ByRef EmailBody As String, ByRef Tribble_ID As String) As String
    Dim TABLE As String
    Dim CICS As String
    Dim PO_NUM As String
    Dim facility As String

    'will need some REGEX checking on these values
    PO_NUM = InputBox("PLEASE ENTER THE PO NUMBER FOR THIS TRIBBLE: ")
    facility = InputBox("PLEASE ENTER THE FACILITY # FOR THIS TRIBBLE: ")
    
    'add CICS to email body
    CICS = AddCICSToEmail(Tribble_ID)
    EmailBody = Replace(EmailBody, "(&CICS)", CICS) 'CICS
    
    'add HMTL table to email body
    TABLE = AddOITableToEmail(PO_NUM, facility, CICS)
    EmailBody = Replace(EmailBody, "(&TABLE)", TABLE) ' TABLE
    EmailBody = Replace(EmailBody, "(&PO)", PO_NUM) 'PO#
    EmailBody = Replace(EmailBody, "(&FACILITY)", facility) 'FACILITY
    
    'ADD FACILITY and PO Number to NOTES file
    Dim SQL As String
    Dim Todays_Date As String
    Todays_Date = CStr(YEAR(Now)) & "-" & CStr(Month(Now)) & "-" & CStr(Day(Now)) & " 00:00:00"
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\Status_Update_New.sql")
    SQL = TribblerSQLStringReplace(SQL, Tribble_ID, , , , , , , , , , , , "Email", , CStr(PO_NUM) & "/" & CStr(facility), Todays_Date, "1")
    InsertDataIntoSnowflake (SQL)
    
    OINREmailBody = EmailBody
    
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 04/16/2021
'PURPOSE: The purpose of this function is to run a query on the the snowflake tribble data set to pull the corresponding CICS
'These CICS are then placed into a list and inserted into a tempalte for an OUTLOOK email
'REFERENCES: QuerySnowFlake
Private Function AddCICSToEmail(ByRef Tribble_ID As String) As String
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim CICS As String
    
    'Segement to pull in the CIC values for the Email Body
    SQL = "SELECT * FROM DW_PRD.TEMP_TABLES.PA_TRIBBLES_CICS WHERE TRIBBLE_ID =" & Tribble_ID
    Set rs = QuerySnowFlake(SQL)
    
    Do While Not rs.EOF
        CICS = CICS & rs.Fields("CICS").Value & ", "
    rs.MoveNext
    Loop
    
    CICS = Left(CICS, Len(CICS) - 2)
    
    AddCICSToEmail = CICS
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 04/16/2021
'PURPOSE: The purpose of this function is to run a specific query for OINR tribble types. This will populate infomration that will be
'placed into an HTML table type and inserted into an OUTLOOK EMAIL
'REFERENCES: DB2.RunQuery, DB2.RunQuery, DB2.RunQuery
Private Function AddOITableToEmail(ByRef PO_NUM As String, ByRef facility As String, ByRef CICS As String) As String
    Dim SQL As String
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\OINR_TABLE.sql")
    SQL = Replace(SQL, "(&PO_NUM)", PO_NUM)
    SQL = Replace(SQL, "(&FACILITY)", facility)
    SQL = Replace(SQL, "(&CICS)", CICS)

    Dim Data As Variant
    Dialog.show "Running query...", "This might take a few seconds"
    Data = DB2.RunQuery(SQL)
    Dialog.Hide
    
    Dim TABLE As String
    TABLE = "<table style='width:50%' border='2'>"
    
    Dim row As Integer
    For row = 0 To arrayLength(Data) - 1
        If row = 0 Then
            TABLE = TABLE + "<tr><th>" + CStr(Data(row, 0)) + "</th><th>" + CStr(Data(row, 1)) + "</th><th>" + CStr(Data(row, 2)) + "</th><th>" + CStr(Data(row, 3)) + "</th><th>" + CStr(Data(row, 4)) + "</th><th>" + CStr(Data(row, 5)) + "</th><th>" + CStr(Data(row, 6)) + "</th></tr>"
        Else
            TABLE = TABLE + "<tr><td>" + CStr(Data(row, 0)) + "</td><td>" + CStr(Data(row, 1)) + "</td><td>" + CStr(Data(row, 2)) + "</td><td>" + CStr(Data(row, 3)) + "</td><td>" + CStr(Data(row, 4)) + "</td><td style = 'background-color:#ffff00;'>" + CStr(Data(row, 5)) + "</td><td style = 'background-color:#ffff00;'>" + CStr(Data(row, 6)) + "</td></tr>"
        End If
    Next row

    TABLE = TABLE + "</table>"
    
    AddOITableToEmail = TABLE
    
End Function
Public Sub VCOTribbleData(OfferNum As String, DIV As String)
    Dim SQL As String
    
    'This query cannot take just a 5 it needs 05
    If DIV = "5" Then
        DIV = CStr("0") & CStr(DIV)
    End If
    
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\MVAUG00_VCO_DATA.sql")
    SQL = Replace(SQL, "(&OFFER_NUM)", OfferNum)
    SQL = Replace(SQL, "(&DIV)", DIV)

    Dim Data As Variant
    Dialog.show "Running query...", "This might take a few seconds"
    Data = DB2.RunQuery(SQL)
    Dialog.Hide
    
    Dim row As Integer
    Dim DataString As String
    For row = 0 To arrayLength(Data) - 1
        With GeneralQueryForm.ListBox1
            .AddItem
            .List(row, 0) = CStr(Data(row, 0))
            .List(row, 1) = CStr(Data(row, 1))
            .List(row, 2) = CStr(Data(row, 2))
            .List(row, 3) = CStr(Data(row, 3))
            .List(row, 4) = CStr(Data(row, 4))
            .List(row, 5) = CStr(Data(row, 5))
            .List(row, 6) = CStr(Data(row, 6))
            .List(row, 7) = CStr(Data(row, 7))
            .List(row, 8) = CStr(Data(row, 8))
            .List(row, 9) = CStr(Data(row, 9))
        End With
    Next row
    
    GeneralQueryForm.Title.Caption = "MVAUG00.VCO"
    GeneralQueryForm.show

End Sub
Public Sub ExportDIV34Data()
    Dim SQL, rngString As String
    Dim rs As ADODB.Recordset
    Dim ws As Excel.Worksheet
    Dim wb As Workbook

    Workbooks.Add.SaveAs Filename:="C:\rtmacros\EXPORT_DIV_34_DATA"
    Workbooks("EXPORT_DIV_34_DATA.xlsx").Activate
    
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\EXPORT_DIV_34_DATA.sql")
    ImportSnowflakeTable SQL
    
    rngString = "A1:" & CStr(Right(Left(Cells(1, 13).Address, 2), 1)) & "1"
    Range(rngString).AutoFilter 'filter on column
    Range(rngString).EntireColumn.AutoFit 'autofit to cells
    
    Workbooks("EXPORT_DIV_34_DATA.xlsx").Save
    Workbooks("EXPORT_DIV_34_DATA.xlsx").Close
    
    MsgBox "EXPORT_DIV_34_DATA.xlsx is saved too C:\rtmacros"
    
End Sub

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 09/17/2020
'Purpose:The purpose of this function is archive the claims into a .csv file automatically.
'This will allow for us to have a backup incase the database ever goes down. Currently this
'Function is set to archive every 50 claims.
Public Sub CheckIfTribbleArchiveIsReady()
    Dim rs As ADODB.Recordset
    Dim ROW_NUM As Integer
    
    Set rs = QuerySnowFlake("SELECT COUNT(*) AS ROW_NUM FROM DW_PRD.TEMP_TABLES.PA_TRIBBLES")
    Do While Not rs.EOF
        ROW_NUM = rs.Fields("ROW_NUM").Value
        rs.MoveNext
    Loop
    
    If ROW_NUM Mod 30 = 0 Then
        Dialog.show "Archiving Process Started", "DO NOT CLOSE OUT OF WINDOWS- WAIT 1 MINUTE AND THEN PRESS OK AFTER PROCESS IS COMPLETED"
        ArchiveSFDataBases "Tribble"
        Dialog.Hide
    End If
    
    
End Sub

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 05/17/2021
'Purpose:The purpose of this function is to take in varaibles from the Status/Notes from in the Tribbler tool. Thes values are then determined if the status has changed
'If the status has changed then a message is autoamtically added to the Notes section. After this function has been completed the Tribbler Status form closes and the
'user is returned to the Tribbler main form
'References:ReadTextFile, TribblerSQLStringReplace, InsertDataIntoSnowflake
Public Sub UpdateTribbleStatusNotes(ByVal TribbleID As String, ByVal OgDate As String, ByVal NumContactMade As String, ByVal OgStatus As String)
    Dim SQL As String
    Dim err As Integer
    Dim Todays_Date As String
    Dim Notes As String
    Todays_Date = CStr(YEAR(Now)) & "-" & CStr(Month(Now)) & "-" & CStr(Day(Now))

    If OgStatus = "Pending" And TribbleStatusUpdateForm.ComboBox1.Value = "Pending" Then
        SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\Status_Update_New.sql")
        SQL = TribblerSQLStringReplace(SQL, TribbleID, , , , , , , , , , , , "Pending", , TribbleStatusUpdateForm.NotesBox.Value, OgDate, CStr(NumContactMade))
    ElseIf OgStatus = "Pending" And TribbleStatusUpdateForm.ComboBox1.Value <> "Closed-Preaudit" Then
        MsgBox "You are unable to update the status of a Pending Tribble here. View Tribble Details to change status"
        SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\Status_Update_New.sql")
        SQL = TribblerSQLStringReplace(SQL, TribbleID, , , , , , , , , , , , "Pending", , TribbleStatusUpdateForm.NotesBox.Value, OgDate, CStr(NumContactMade))
    ElseIf OgStatus <> TribbleStatusUpdateForm.ComboBox1.Value Then
        Notes = TribbleStatusUpdateForm.NotesBox.Value & vbNewLine & "Status Changed From " & OgStatus & " To " & TribbleStatusUpdateForm.ComboBox1.Value & " BY:" & Environ("Username") & " ON " & Todays_Date & vbNewLine
        SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\Status_Update_New.sql")
        SQL = TribblerSQLStringReplace(SQL, TribbleID, , , , , , , , , , , , TribbleStatusUpdateForm.ComboBox1.Value, , Notes, OgDate, CStr(NumContactMade))
    Else
        SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\Status_Update_New.sql")
        SQL = TribblerSQLStringReplace(SQL, TribbleID, , , , , , , , , , , , TribbleStatusUpdateForm.ComboBox1.Value, , TribbleStatusUpdateForm.NotesBox.Value, OgDate, CStr(NumContactMade))
    End If
       
    err = InsertDataIntoSnowflake(SQL)
    If err = 0 Then
        MsgBox "Tribble Status/Notes Successfully Updated"
        TribbleStatusUpdateForm.ContactMadeCheckBox.Value = False
        TribbleStatusUpdateForm.BackDateCheckBox.Value = False
        TribbleStatusUpdateForm.Hide
    Else
        MsgBox "An Error occured Updating Status/Notes. Please contact BA"
    End If
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 05/17/2021
'Purpose: The purpose of this function is to fill in the ASM field in the New Tribble FORM
Public Sub FillASMFormField(ByVal DIV As String)
    'check to see if the fields needed are filled in
    If TribbleFormNew.CicsListBox.ListCount = 0 Or DIV = vbNullString Then
        MsgBox "Please enter in a valid CIC and DIV First"
        Exit Sub
    End If
    
    'Clear ifanything is currently in the ASM box
    TribbleFormNew.ASMTextBox.Value = ""
    
    'check if multiple divisions
    If Contains(DIV, ",") = True Then
        Dim SplitDiv() As String
        SplitDiv = split(DIV, ",")
        DIV = SplitDiv(0)
    End If
    
    Dim CIC, PROD_CD, CTGRY_CD  As String 'pull first CIC value
    CIC = TribbleFormNew.CicsListBox.List(0)
    
    If Len(CIC) = 7 Then
        PROD_CD = Left(CIC, 1)            '1st digit of CIC
        CTGRY_CD = Right(Left(CIC, 3), 2) '2nd and 3rd digits of CIC
    Else
        PROD_CD = Left(CIC, 2) 'first 2 digits of CIC 30010010 DIV 32
        CTGRY_CD = Right(Left(CIC, 4), 2) '3rd and 4th digits of CIC
    End If
    
    Dim SQL As String 'Run the query
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\ASM_FILL.txt")
    SQL = Replace(SQL, "(&PROD_CD)", PROD_CD)
    SQL = Replace(SQL, "(&CTGRY_CD)", CTGRY_CD)
    SQL = Replace(SQL, "(&DIV)", DIV)
    
    Dim Data As Variant
    Dialog.show "Running query...", "This might take a few seconds"
    Data = DB2.RunQuery(SQL)
    Dialog.Hide
    
    Dim item As Variant
    'populate asm box
    For Each item In Data
        If item <> "ASM" Then
            TribbleFormNew.ASMTextBox.Value = item
            Exit For
        End If
    Next item
    
    'Error message if no ASM is found
    If TribbleFormNew.ASMTextBox.Value = vbNullString Then
        MsgBox "ASM could not be found"
    End If

End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 06/01/2021
'Purpose:Version2 of Tribbler [IN TESTING]
'References:QuerySnowFlake
Public Sub AddTribblesToListV2(ByVal SQL As String)
On Error Resume Next
    Dim rs As ADODB.Recordset
    Dim item As Variant
    Dim counter As Integer
    
    TribbleFormV2.TribbleListBox.Clear

    Set rs = QuerySnowFlake(SQL)

    Dialog.show "Populating Data...", "This might take a few seconds"
    
    If IsRecordsetEmpty(rs) Then
        MsgBox ("No Tribbles Exist")
        Dialog.Hide
        Exit Sub
    End If
    
    counter = 0
    With TribbleFormV2.TribbleListBox
        With TribbleFormV2.HeaderListBox
            .AddItem
            .List(0, 0) = "TribbleID"
            .List(0, 1) = "Contact Made"
            .List(0, 2) = "Issue"
            .List(0, 3) = "Status"
            .List(0, 4) = "Division"
            .List(0, 5) = "FirstCIC"
            .List(0, 6) = "ASM"
            .List(0, 7) = "AUDITOR"
            .List(0, 8) = "OFFERNUM"
            .List(0, 9) = "VendorName"

        End With
        
        Do While Not rs.EOF
            .AddItem
            .List(counter, 0) = rs.Fields("TRIBBLE_ID").Value
            .List(counter, 1) = rs.Fields("LAST_CONTACT_MADE").Value
            .List(counter, 2) = rs.Fields("ISSUE").Value
            .List(counter, 3) = rs.Fields("STATUS").Value
            .List(counter, 4) = rs.Fields("DIVISION").Value
            .List(counter, 5) = rs.Fields("FIRSTCIC").Value
            .List(counter, 6) = rs.Fields("ASM").Value
            .List(counter, 7) = rs.Fields("PREAUDITOR").Value
            .List(counter, 8) = rs.Fields("OFFER_NUM").Value
            .List(counter, 9) = rs.Fields("VENDOR_NAME").Value

            counter = counter + 1
            rs.MoveNext
        Loop
    End With
    Dialog.Hide
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 01/20/2021
'Purpose: The purpose of this function is a helper function to check to see if a cic value
'exists in a form before adding it to the list. It returns a boolean
'References:None
Public Function CheckIfCICValueExistsInListBox(ByVal InputVal As String) As Boolean
    Dim index As Integer

    For index = 0 To TribbleFormNew.CicsListBox.ListCount - 1
        If TribbleFormNew.CicsListBox.List(index) = InputVal Then
            CheckIfCICValueExistsInListBox = True
            Exit Function
        End If
    Next index
    CheckIfCICValueExistsInListBox = False
End Function
'////////////////////////////////////////////////TRIBBLER NOTES AND STATUS/////////////////////////////////////////////////////
'Created: 10/16/2020
'Purpose: The purpose of this function is to initlize the Tribble FORM Notes section
'References:
Public Sub Initalize_Notes(ByRef TribbleID As String)
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim ColumnHeaders As Variant
    
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\New_Get_Tribble.sql")
    SQL = Replace(SQL, "(&Tribble_ID)", TribbleID)
    
    Set rs = QuerySnowFlake(SQL)
    If IsRecordsetEmpty(rs) Then
        Exit Sub
    End If
    
    Do While Not rs.EOF
        TribbleStatusUpdateFormV2.StatusComboBox.Value = rs.Fields("STATUS").Value
        TribbleStatusUpdateFormV2.NotesBox.Value = rs.Fields("TRIBBLE_NOTES").Value
        TribbleStatusUpdateFormV2.EmailTextBox1.Value = rs.Fields("NUM_CONTACT_MADE").Value
        TribbleStatusUpdateFormV2.LastContactMadeTextBox.Value = rs.Fields("LAST_CONTACT_MADE").Value
    rs.MoveNext
    Loop
    
End Sub
''Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 10/28/2021
'Purpose:The purpose of this function is to check if a tribble exists with the same CIC and division. This is a helper function.
'IF True then another Tribble Does Exist. If False then another Tribble Does not Exists
Public Function CheckIfTribbleExistsV2(ByRef DIV As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim SQL As String
    
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\New_Tribble_Check_Exists_DIV_CICS.sql")
    SQL = Replace(SQL, "(&DIVISION)", DIV)
    SQL = Replace(SQL, "(&CICS)", ProduceTribbleCicsString)
    
    Set rs = QuerySnowFlake(SQL)
    
    If IsRecordsetEmpty(rs) Then
         CheckIfTribbleExistsV2 = False
    Else
        CheckIfTribbleExistsV2 = True
    End If
    
End Function

''Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 10/28/2021
'Purpose:The purpose of this function is to check if a tribble exists with the same CIC and division. This is a helper function.
'IF True then another Tribble Does Exist. If False then another Tribble Does not Exists
Public Function CheckIfTribbleExistsRecordSet(ByRef DIV As String) As ADODB.Recordset
    Dim SQL As String
    
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\New_Tribble_Check_Exists_DIV_CICS.sql")
    SQL = Replace(SQL, "(&DIVISION)", DIV)
    SQL = Replace(SQL, "(&CICS)", ProduceTribbleCicsString)
    
    Set CheckIfTribbleExistsRecordSet = QuerySnowFlake(SQL)

End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 11/05/2021
'Purpose: The purpose of this function is to reproduce the email for the tribble by using the Tribble ID
'References: Produce Tribble Email, querysnowflake
Public Sub ReProduceTribbleEmail(ByVal TribbleID As String)
    Dim SQL As String
    Dim rs As ADODB.Recordset
    SQL = "SELECT * FROM DW_PRD.TEMP_TABLES.PA_TRIBBLES WHERE TRIBBLE_ID = '" & TribbleID & "'"
    Set rs = QuerySnowFlake(SQL)
    
    If IsRecordsetEmpty(rs) Then
        MsgBox "Theis Tribble's email could not be popualted"
    Else
        ProduceTribbleEmail rs.Fields("DIVISION").Value, rs.Fields("WIMS_VNDR_NUM").Value, rs.Fields("ISSUE").Value, rs.Fields("VENDOR_NAME").Value, TribbleID, rs.Fields("OFFER_NUM").Value, rs.Fields("ASM").Value, rs.Fields("FIRSTCIC").Value
    End If
        
End Sub
'-------------------------------------------------------------------------------------
'----------------------------------------TRIBBLE_FORM_NEW_V2--------------------------
'-------------------------------------------------------------------------------------
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 09/23/2020
'Purpose: This checks for empty values in the form
'References:None
'TODO need to implement the abilit yto check to see if a date is larger than another date
Private Function ValidateNewTribbleForm(WIMS As String, VENDOR_NAME As String, ASM As String, issue As String, SDATE As String, EDATE As String) As Boolean
    
    If TribbleFormNewV2.CicsListBox.ListCount = 0 Or WIMS = "" Or VENDOR_NAME = "" Or ASM = "" Or issue = "" Or _
    IsDate(SDATE) = False Or IsDate(EDATE) = False Or IsNumeric(WIMS) = False Or (Left$(WIMS, 1) = "0") = False Then
        ValidateNewTribbleForm = False
    Else
        ValidateNewTribbleForm = True
    End If
    
End Function

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 09/23/2020
'Purpose:The purpose of this function is check if a Tribble exists with the same CIC and Division value
'References: Only check for tribble by the DIV and CICS
Private Function CheckNewTribbleExists(DIV) As Boolean
    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim item As Variant
    Dim MsgBoxString As String
    Dim answer As String
    
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\New_Tribble_Check_Exists_DIV_CICS.sql")
    SQL = Replace(SQL, "(&DIVISION)", DIV)
    SQL = Replace(SQL, "(&CICS)", ProduceTribbleCicsStringV2)

    Set rs = QuerySnowFlake(SQL)

    If IsRecordsetEmpty(rs) Then
        CheckNewTribbleExists = True
    Else
        CheckNewTribbleExists = False
    End If

End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 09/08/2020
'Purpose: 'The division check boxes are filled in with the information from the
'References:
Public Sub FillDivisionBoxes(Division As String)
    Dim MyResult() As String
    Dim i As Integer
    
    Select Case True
        Case Division = "All":  TribbleFormNewV2.CheckBox15.Value = True
        Case Division = "5":    TribbleFormNewV2.CheckBox2.Value = True
        Case Division = "05":   TribbleFormNewV2.CheckBox2.Value = True
        Case Division = "15":   TribbleFormNewV2.CheckBox3.Value = True
        Case Division = "19":   TribbleFormNewV2.CheckBox4.Value = True
        Case Division = "17":   TribbleFormNewV2.CheckBox1.Value = True
        Case Division = "20":   TribbleFormNewV2.CheckBox8.Value = True
        Case Division = "24":   TribbleFormNewV2.CheckBox18.Value = True
        Case Division = "25":   TribbleFormNewV2.CheckBox7.Value = True
        Case Division = "27":   TribbleFormNewV2.CheckBox6.Value = True
        Case Division = "28":   TribbleFormNewV2.CheckBox5.Value = True
        Case Division = "29":   TribbleFormNewV2.CheckBox10.Value = True
        Case Division = "30":   TribbleFormNewV2.CheckBox9.Value = True
        Case Division = "32":   TribbleFormNewV2.CheckBox12.Value = True
        Case Division = "33":   TribbleFormNewV2.CheckBox11.Value = True
        Case Division = "34":   TribbleFormNewV2.CheckBox13.Value = True
        Case Division = "35":   TribbleFormNewV2.CheckBox17.Value = True
        Case Division = "65":   TribbleFormNewV2.CheckBox16.Value = True
    End Select
    
End Sub

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/27/2020
'Purpose: Add CIC vlaues to DB.
'References:When updating a value or adding a new one
Private Sub AddCICValuesToDB(Tribble_ID As String)
    Dim i As Integer
    Dim SQL As String
    Dim queryvalues As String
    Dim err As Integer
    
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\New_CICS.sql")
    
    For i = 0 To TribbleFormNewV2.CicsListBox.ListCount - 1
        If i = (TribbleFormNewV2.CicsListBox.ListCount - 1) Then
            queryvalues = queryvalues & "(" & Tribble_ID & "," & TribbleFormNewV2.CicsListBox.List(i) & ");"
        Else
            queryvalues = queryvalues & "(" & Tribble_ID & "," & TribbleFormNewV2.CicsListBox.List(i) & "), "
        End If
    Next i
    
    SQL = SQL & queryvalues
    
    err = InsertDataIntoSnowflake(SQL)
    If err <> 0 Then
        MsgBox "And Error Occured when Adding CIC Values to DB. Please contact BA"
    End If
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/27/2020
'Purpose:Purpose is to fill the CCI box at the opening of the form.
'References: Access connection for getting cics
Public Sub FillCICBox(Tribble_ID As String)
    Dim rs As ADODB.Recordset
    Dim item As Variant
    Dim SQL As String

    CicsListBox.Clear                                               'ensure that box is clear before opening

    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\New_Get_CICS.sql")
    SQL = Replace(SQL, "(&Tribble_ID)", Tribble_ID)

    Set rs = QuerySnowFlake(SQL)
    If IsRecordsetEmpty(rs) Then
        TribbleFormNew.show                       'open form after values are populated 'TODO DELETE ME
        'Exit Sub
    End If

    Do While Not rs.EOF
        CicsListBox.AddItem rs.Fields("CICS").Value
    rs.MoveNext
    Loop

    TribbleFormNew.show                       'open form after values are populated

End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/27/2020
'Purpose: This produces a query to insert a new tribble into the databse
'References: Access connection and AddCICValuesToDB, AddCICValuesToDBv FindTribbleId, ProduceTribbleEmail
Private Sub NewTribbleInsertQuery(ByRef Tribble_ID As String, ByRef DIV As String, ByRef WIMS As String, ByRef ASM As String, _
ByRef VENDOR_NAME As String, ByRef Offer_Num As String, ByRef TimeFrame_Start As String, ByRef TimeFrame_End As String, ByRef issue As String, ByRef Status As String, _
ByRef TRIBBLER As String, ByRef Tribble_Notes As String, ByRef PREAUDITOR As String, ByRef FIRSTCIC As String)
    Dim SQL As String
    Dim Todays_Date As String
    Dim err As Integer
    
    'Define date and Tribble ID
    Todays_Date = CStr(YEAR(Now)) & "-" & CStr(Month(Now)) & "-" & CStr(Day(Now)) & " 00:00:00"
    Tribble_ID = CStr(CInt(GetLastTribbleID) + 1)
    
    'Creation of SQL Query
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\New_Tribble.sql")
    SQL = TribblerSQLStringReplace(SQL, Tribble_ID, Todays_Date, DIV, WIMS, ASM, FIRSTCIC, UCase(PREAUDITOR), _
    VENDOR_NAME, Offer_Num, TimeFrame_Start, TimeFrame_End, issue, Status, UCase(TRIBBLER), Tribble_Notes, Todays_Date, 0)

    err = InsertDataIntoSnowflake(SQL)
    If err = 0 Then
        MsgBox "Tribble Added Successfully to Database"
        AddCICValuesToDB Tribble_ID                     'add CICS to DB using tribble ID
        ProduceTribbleEmail DIV, WIMS, issue, VENDOR_NAME, Tribble_ID, Offer_Num, ASM, FIRSTCIC
    Else
        MsgBox "An Error Occured when adding the Tribble. Please Contact BA"
    End If
       
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 08/27/2020
'Purpose: The purpose is to update the existing tribbles in the database. It first updates
'the tribbles in the 2020_calendar and then it deletes the old CICs and then adds the new CICs.
'References: access connection, delete old cics, and add cicvalues to db
'Private Sub UpdateExistingTribbleDB(ByRef Tribble_ID As OString, ByRef DIV As String, ByRef WIMS As String, ByRef ASM As String, _
'ByRef VENDOR_NAME As String, ByRef Offer_Num As String, ByRef TimeFrame_Start As String, ByRef TimeFrame_End As String, ByRef issue As String, ByRef Status As String, _
'ByRef TRIBBLER As String, ByRef Tribble_Notes As String, ByRef PREAUDITOR As String, ByRef FIRSTCIC As String)
'    Dim sql As String
'    Dim err As Integer
'
'    'Creation of SQL Query
'    sql = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\New_Update_Tribble.sql")
'    sql = TribblerSQLStringReplace(sql, Tribble_ID, , DIV, WIMS, ASM, , , VENDOR_NAME, Offer_Num, TimeFrame_Start, TimeFrame_End, issue, Status, UCase(TRIBBLER), Tribble_Notes, , "")
'
'    err = InsertDataIntoSnowflake(sql)
'    If err = 0 Then
'        MsgBox "Tribble Successfully Updated"
'        DeleteOldCics Tribble_ID        'delete old CICs
'        AddCICValuesToDB Tribble_ID     'add CICs to DB using tribble ID
'    Else
'        MsgBox "An Error occured updating the Tribble. Please Contact BA"
'    End If
'
'End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 06/14/2021
'Purpose: Called by clicking submit/update in Tribbler Tool. If command is INSERT then add If command is UPDATE than update
Public Sub INSERT_UPDATE_TRIBBLE(Command As String, Tribble_ID As String, DIV As String, WIMS_Vndr_Num As String, ASM As String, _
VENDOR_NAME As String, Offer_Num As String, TimeFrame_Start As String, TimeFrame_End As String, issue As String, Status As String, _
TRIBBLER As String, Tribble_Notes As String, PREAUDITOR As String, FIRSTCIC As String)
On Error GoTo err
    'Validate Form
    If VALIDATE_TRIBBLE_DATA(WIMS_Vndr_Num, VENDOR_NAME, ASM, issue, TimeFrame_Start, TimeFrame_End, DIV) = False Then GoTo valcatch

    'Insert Data into Tribble_CICS  & Tribble Table
    If Command = "UPDATE" Then
        UpdateExistingTribbleDB Tribble_ID, DIV, WIMS_Vndr_Num, ASM, VENDOR_NAME, Offer_Num, TimeFrame_Start, TimeFrame_End, issue, Status, TRIBBLER, Tribble_Notes, PREAUDITOR, FIRSTCIC
    Else
        NewTribbleInsertQuery Tribble_ID, DIV, WIMS_Vndr_Num, ASM, VENDOR_NAME, Offer_Num, TimeFrame_Start, TimeFrame_End, issue, Status, TRIBBLER, Tribble_Notes, PREAUDITOR, FIRSTCIC
    End If
    
    'SUCCESS
    Exit Sub
    
err:
    MsgBox "Error Occured During INSERT_UPDATE_TRIBBLE: Contact BA"
    Exit Sub
    
valcatch:
    MsgBox "FORMATTING ERROR"
    Exit Sub
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 06/14/2021
'Purpose:
Public Function VALIDATE_TRIBBLE_DATA(ByRef WIMS As String, ByRef VENDOR_NAME As String, ByRef ASM As String, ByRef issue As String, ByRef SDATE As String, ByRef EDATE As String, ByRef DIV As String) As Boolean
    'Validate TEXT Boxes
    VALIDATE_TRIBBLE_DATA = ValidateNewTribbleForm(WIMS, VENDOR_NAME, ASM, issue, SDATE, EDATE)
    'Check if Tribble exists
    VALIDATE_TRIBBLE_DATA = CheckNewTribbleExists(DIV)
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 06/14/2021
'Purpose:
Public Sub AutoFiller(OfferNum As Long)
    On Error GoTo catch
    Dim Data As Variant
    Dim SQL As String
    
    SQL = ReadTextFile(ENV.use("SQLTRIBBLERFOLDERPATH") & "\Get_Nopa_Info_V2.sql")
    SQL = Replace(SQL, "(&Offer_Num)", OfferNum)
    
    Data = DB2.RunQuery(SQL)
    
    Dim Search As String
    Dim item As Variant
    
    For Each item In Data
        Select Case True
            Case item = "RTL_ITM_NBR":      Search = "CICS"
            Case item = "PRFRM_START_DT":   Search = "StartDate"
            Case item = "PRFRM_END_DT":     Search = "EndDate"
            Case item = "DIV":              Search = "Div"
            Case item = "ASM":              Search = "ASM"
            Case item = "VEND_NUM":         Search = "WIMS_VEND_NUM"
            Case item = "VENDOR_NAME":      Search = "VendNM"
            Case Search = "CICS":           TribbleFormNewV2.CicsListBox.AddItem item
            Case Search = "StartDate":      TribbleFormNewV2.StartTextBox.Value = item
            Case Search = "EndDate":        TribbleFormNewV2.EndTextBox.Value = item
            Case Search = "VendNM":         VendorNameClearExtraSpaces (item)
            Case Search = "WIMS_VEND_NUM":  TribbleFormNewV2.WimsTextBox.Value = item
            Case Search = "ASM":            TribbleFormNewV2.ASMTextBox.Value = item
            Case Search = "Div":            FillDivisionBoxes (CStr(item))
        End Select
    Next item
    
    Exit Sub
'Error Handling
catch:
    If err.Number <> 0 Then
        CatchErrorController "Tribbler.AutoFillerTesting"
        Exit Sub
    End If
    
End Sub
'Used when populating a new Tribble in TribbleFormNew
'New Tribbles have their CICS and divison compared.
'This turns the list of CICS in the list box into a String (1111,2222,3333).
Public Function ProduceTribbleCicsStringV2() As String
    On Error GoTo catch
    Dim i As Integer
    
    ProduceTribbleCicsStringV2 = "("
    For i = 0 To TribbleFormNewV2.CicsListBox.ListCount - 1
        If i = (TribbleFormNewV2.CicsListBox.ListCount - 1) Then
            ProduceTribbleCicsStringV2 = ProduceTribbleCicsStringV2 & TribbleFormNewV2.CicsListBox.List(i)
        Else
            ProduceTribbleCicsStringV2 = ProduceTribbleCicsStringV2 & TribbleFormNewV2.CicsListBox.List(i) & ", "
        End If
    Next i
    ProduceTribbleCicsStringV2 = ProduceTribbleCicsStringV2 & ")"
    
    Exit Function
'Error Handling
catch:
    If err.Number <> 0 Then
        CatchErrorController "Tribbler.ProduceTribbleCicsStringV2"
        Exit Function
    End If
    
End Function

