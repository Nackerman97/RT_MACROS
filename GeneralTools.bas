Attribute VB_Name = "GeneralTools"
Option Explicit
Option Compare Text

Public MouseStart
Public LastMoved

'@AUTHOR: ROBERT TODAR

'DEPENDENCIES
' -

'PUBLIC FUNCTIONS
' -
' -
' -
' -

'PRIVATE METHODS/FUNCTIONS
' -
' -
' -
' -

'NOTES:
' - THIS WAS A GENERAL PLACE TO STORE CODE, THINGS THAT WERE KINDA ISOLATED AND NOT RELATED
'   TO ANYTHING ELSE.
' - OPENS UP USERFORMS/

'TODO:
' - CLEAN ENTIRE MODULE UP. CURRENTLY THINGS ARE NOT LABELED WELL. NOT SURE IF ALL
' - ADD MORE NOTES!!
' - REMOVE UNEEDED METHODS.
' - NEED TO ADD ALL FUNCTIONS TO THIS LIST

' CALLS RT_MACROS_BOX USERFORM
Public Sub RT_Macros_Box()
Attribute RT_Macros_Box.VB_ProcData.VB_Invoke_Func = "e\n14"
    On Error GoTo catch
    DisplayGridForm
    
    Exit Sub
catch:
    On Error Resume Next
    Workbooks.Add
    DisplayGridForm
End Sub

' CALLS BOUNCER QUERY _BOX USERFORM
Public Sub BouncerQueryBox()
    PromptForBouncerReportEmailer
    LogCode "Bouncer Report Emailer"
End Sub

Public Sub findingsReport()
    findings.show
End Sub

' CALLS biller USERFORM
Public Sub BillerTesting()
    BillerForm.show
End Sub

' TESTING AN OI CASE AUTO BILLER
Public Sub VCPivotBillingTest()
    BillerForm.show
End Sub

' OPEN TOOLS CATALOG
Public Sub ToolsCatalog()
    OpenAnyFile "K:\AA\SHARE\AuditTools\rtmacros\notes\index.html"
End Sub

' MAKES 20/1 DATA FROM CICS
Public Function CicAllowForAllDiv() As Boolean
    Dim CICList As String
    
    CICList = InputBox("What CICS?")
    If CICList = "" Then Exit Function
    
    Call CreateAllowanceOffersWorkbook("", CICList)
End Function
'Added by Nick so that this function can be found in the quick access toolbar
'01/19/2021. It needs to be a sub to be added
Public Sub RunBouncerLookupQuery()
    BouncerLookupQuery
End Sub
' CALLS MY TOOLS USERFORM
Public Sub MyToolsBox()
Attribute MyToolsBox.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim ac As Worksheet
    
    On Error GoTo catch
    Set ac = ActiveSheet
    
    'Testing for new form
    If ac.name = "20-1" Then
        FillAllowanceFromOffer
    ElseIf ac.name = "VC" Then
        MessageBox "Attempting to navigate PACS to the EDI screen"
        CasePacsLookup
        
    ElseIf ac.name = "VC Table" Or ac.name = "VC Pivot" Then
        MessageBox "Bouncer Bill Fill"
        OIBouncerFill
        
    ElseIf ac.name = "Bouncer list" Then
        searchForPo
        
    ElseIf Left(ac.name, Len("Placement")) = "Placement" Then
        GotoPf22
    End If
    
    MessageBox CloseForm:=True
catch:
End Sub

Private Sub GotoPf22()
    Dim DIV As String
    DIV = Cells(ActiveCell.row, FindColumnData(ActiveSheet, "DIVISION").column).Value
    
    Dim Warehouse As String
    Warehouse = Right(Cells(ActiveCell.row, FindColumnData(ActiveSheet, "FACILITY").column).Value, 2)
    
    Dim CIC As String
    CIC = Cells(ActiveCell.row, FindColumnData(ActiveSheet, "CORP_ITEM_CD").column).Value
    
    Dim bz As New blueZoneObject
    
    If bz.ConnectToActiveSession(True) = False Then
        Exit Sub
    End If
    
    'CLEAR ANY HANG UPS
    bz.Sendkeys ("<RESET>")
    
    'CLEAR OUT OF ANY CURRENT SCREEN
    If bz.ClearOutToABlankScreen = False Then
        
        'RETURN ERROR MESSAGE
        Exit Sub
        
    End If
   
    'FROM BLANK SCREEN, NAVIGATE TO PACS
    bz.Sendkeys ("PACS<ENTER>")
    
    'ONCE IN PACS, NAVIGATE TO PROPER DIVISION
    If bz.WaitForStringOnPage("*PACS -  Logo*Division:*") = False Then
        Exit Sub
    End If
    
    bz.PutString DIV, 3, 18
    bz.Sendkeys "<Enter>"
    
    'ONCE IN DIVISION, SWITCH PF SET 1 TO PF SET 2, THEN ENTER THE 17/2
    bz.WaitForStringOnPage "*PF*SET:*1*"
    bz.Sendkeys "<PF19>"
    bz.Sendkeys "<Enter>"
    
    'put cic info
    bz.PutString Warehouse, 3, 13
    bz.PutString CIC, 3, 24
    bz.Sendkeys "<Enter>"

    
    'goto pf22
    bz.Sendkeys "<PF22>"
    bz.Sendkeys "<Enter>"
    AppActivate "Allowance Billing"
End Sub

Public Function DeveloperAnalytics()
    On Error Resume Next
    Application.EnableEvents = True
    Workbooks.Open "K:\AA\SHARE\AuditTools\rtmacros\codelog.xlsm", , True, , , , True
End Function

Public Function BouncerReport()
    On Error Resume Next
    Application.EnableEvents = True
    Workbooks.Open "K:\AA\SHARE\AuditTools\rtmacros\bouncers.xlsm", , True, , , , True
End Function

'CREATED FOR KRISTINA - WILL ALLOW HER TO TOGGLE FILTER ON AND OFF
Public Sub coolBeansFilter()
    Range("A1").AutoFilter
End Sub

' CALLS BOUNCER STATUS USERFORM
Public Sub BouncerStatusForm()
    On Error Resume Next
    Workbooks("bouncer stats.xls").Activate
    BouncerStats.show
End Sub

' CALLS CASEAUDITFORM USERFORM
Public Sub CaseVcOfferBox()
    LogCode "VC Offer"
    CaseAuditForm.show
End Sub

' CALLS J4U USERFORM
Public Sub j4uBox()
    On Error Resume Next
    J4UBiller.show
End Sub

' CREATED FOR LINDSAY. FORMATS UPCS INTO ONE STRING, WITHOUT THE LINEBREAKS
Public Sub RemoveLinesFromTextFile()
    'DECLARE
    Dim fso As cFileSystemObject
    Dim ts As Object
    Dim fd As FileDialog
    Dim ActionClicked As Boolean
    Dim s As String

    'INITIAL SET
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.InitialFileName = Environ("UserProfile") & "/desktop/"
    
    'OPEN FILE EXPLORE TO SELECT TEXT FILE
    fd.AllowMultiSelect = False
    ActionClicked = fd.show

    If ActionClicked Then
        Set fso = New cFileSystemObject
        Set ts = fso.OpenTextFile(fd.SelectedItems(1), ForReading)
        
        Do Until ts.AtEndOfStream
            s = s & ts.ReadLine
        Loop
        
        ClipboardSet s
        MsgBox "Data copied to the clipboard!"
        
        ts.Close
    End If
End Sub

' FILLS ASM INFORMATION ON WITHTRIBBLE WORKBOOK. ASM DATA STORED IN THISWORKBOOK
Public Sub ASMFill(Optional Bo As Boolean)
        If Not ActiveWorkbook.name Like "*Tribble*" Then
            MsgBox "Active Workbook is """ & ActiveWorkbook.name & """." & _
            vbNewLine & "Must have ""With Tribble.xlsx"" as the active workbook."
            Exit Sub
        End If
        Application.ScreenUpdating = False
        
        Dim wb As Workbook
        Dim ac As Range
        Dim ans As VbMsgBoxResult
        
        Dim SMIC As String
        Dim DIV As String
        Dim Val As String
        Dim RowCount As Integer
        
        Set ac = Range("D" & ActiveCell.row)
        
        SMIC = ac.Offset(0, 1).Value
        DIV = ac.Offset(0, -2).Value
        
        If Len(SMIC) < 6 Or Len(DIV) < 1 Then: GoTo CantFind
        
        If DIV = "5" Then DIV = "05"
        
        'Picker for Texas Divsions
        If DIV = "20" Then
            ans = MsgBox("Yes for: 20: Southern,    No for:  23: Houston", vbYesNo)
            If ans = vbCancel Then: Exit Sub
            If ans = vbYes Then: DIV = "20"
            If ans = vbNo Then: DIV = "23"
        End If
        
        If Len(SMIC) = 7 Then
            SMIC = "0" & SMIC
        End If
        
        'IF FOUND, THEN SQL ASM LIST DATA TO WORKSHEET
        Set wb = Workbooks.Add
        SQL_AsmList wb.ActiveSheet
        
        
        Val = Left(SMIC, 4) & DIV
        On Error GoTo CantFind
        RowCount = wb.ActiveSheet.Range("A1", wb.ActiveSheet.Range("A" & Rows.Count).End(xlUp)).Count
        
        wb.ActiveSheet.Range("J2").Value = "=CONCATENATE(" & wb.ActiveSheet.Range("I2").Address(False, True) & "," & wb.ActiveSheet.Range("D2").Address(False, True) & ")"
        wb.ActiveSheet.Range("J2").AutoFill (wb.ActiveSheet.Range("J2", wb.ActiveSheet.Range("J" & RowCount)))
        
        
        ac.Value = WorksheetFunction.Proper(wb.ActiveSheet.Range("J:J").Find(what:=Val, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Offset(0, -3).Value)
        
        wb.Close False
        Application.ScreenUpdating = True
        Exit Sub
CantFind:
        Application.ScreenUpdating = True
        MsgBox "Can't find ASM, Check Div or CIC"
End Sub

' ADDS MY OWN STYLE FOR TABLES
Public Function customTableStyle() As String
    Dim TbleName As String
    
    TbleName = "Robert's Style"
    customTableStyle = TbleName
    
    On Error GoTo AlreadyExists
    
    'ADD NEW TABLE FORMAT
    ActiveWorkbook.TableStyles.Add (TbleName)
    With ActiveWorkbook.TableStyles(TbleName)
        .ShowAsAvailablePivotTableStyle = False
        .ShowAsAvailableTableStyle = True
        .ShowAsAvailableSlicerStyle = False
        .ShowAsAvailableTimelineStyle = False
    End With
    
    'header
    With ActiveWorkbook.TableStyles(TbleName).TableStyleElements(xlHeaderRow). _
        Interior
        .pattern = xlSolid
        .PatternColor = 10855845
        .Color = RGB(165, 165, 165)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With ActiveWorkbook.TableStyles(TbleName).TableStyleElements(xlHeaderRow). _
        Font
        .Bold = True
        .Color = vbWhite
    End With
    
    'first row stripe
    ActiveWorkbook.TableStyles(TbleName).TableStyleElements(xlRowStripe1).Clear
    With ActiveWorkbook.TableStyles(TbleName).TableStyleElements(xlRowStripe1). _
        Interior
        .pattern = xlSolid
        .Color = 15592941
        .TintAndShade = 0
        .PatternTintAndShade = 0.799890133365886
    End With
    With ActiveWorkbook.TableStyles(TbleName).TableStyleElements(xlRowStripe1). _
        Borders(xlEdgeTop)
        .Color = 13224393
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveWorkbook.TableStyles(TbleName).TableStyleElements(xlRowStripe1). _
        Borders(xlEdgeBottom)
        .Color = 13224393
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
AlreadyExists:
End Function

' CREATES LIST.TXT ON DESKTOP OF UPC #'S FROM DEAL DETAIL SHEET
Public Sub CreateList_txt(Optional Bo As Boolean)
    If Not ActiveSheet.name Like "*Deal Detail*" Then
        CreateListForm.show
        Exit Sub
    End If
    
    'DECLARE VARIABLES
    Dim fso As Object
    Dim ts As Object
    Dim myCollection As New Collection
    
    Dim r As Range
    Dim c As Range
    Dim s As String
    Dim L As Long
    
    'INITIAL SET (GETS UPC RANGE FROM DEAL DETAIL PAGE, CREATES/OPENS TXT FILE)
    On Error GoTo err '****
        Set r = FindHeading(ActiveSheet, "UPCCase").Offset(1)
        Set r = Range(r, r.Offset(-1, 1).End(xlDown))
    On Error Resume Next '****
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(Environ("userprofile") & "\desktop\list.txt", 2, True)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'FOR EACH - GOES THROUGH RANGE OF UPC #'S
    'SETS THE PROPER FORMAT FOR THEM
    'THEN ADDS UPC TO COLLECTION IF IT IS UNIQUE UPC #
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each c In r
        
        'CREATES NEEDED LENGTH OF UPC #
        s = Left(c, Len(c) - 1)
        
        If Len(s) < 13 Then
            Do Until Len(s) = 13
                s = "0" & s
            Loop
        End If
        
        'CHECKS IF UPC IS ALREADY IN THE LIST, ADDS TO COLLECTION OR MOVES ON DEPENDING
        If myCollection.Count > 0 Then
            For L = 1 To myCollection.Count
                If myCollection(L) = s Then: GoTo Nxt
            Next L
        End If
        
        myCollection.Add s
Nxt:
    Next c
        
        ''''''''''''''''''''''''''''''''''''''''''''''
        'INPUT NUMBERS TO LIST.TXT FILE ON DESKTOP
        'ADDING COMMAS AND LINE BREAKS.
        ''''''''''''''''''''''''''''''''''''''''''''''
        For L = 1 To myCollection.Count
            ts.Write myCollection(L)
            
            If L < myCollection.Count Then
                ts.Write ","
                ts.WriteLine
            End If
            
        Next L
    
    'BEST PRACTICES
    ts.Close
    Set fso = Nothing
    Set myCollection = Nothing

    MessageBox "Created List.txt on your desktop", True
    'MyMsgboxForm.Show: MyMsgboxForm.Label1.Caption = "Created List.txt on your desktop": MyMsgboxForm.DisplayOnlyForAMoment
Exit Sub
'''''''''''''''''
'ERROR HANDLING
'''''''''''''''''
err: '****
    MsgBox "Deal Detail sheet not active"
End Sub

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 09/30/2020
'Purpose:This function takes the filelocation and name as string.
'Must be a .txt file.
'It breaks up the file into a string that can be passed to an email function
'Used by Tribbler and Placement
Public Function FormatTemplateEmail(Filename) As String
    On Error GoTo catch
    Dim FileNum As Integer
    Dim DataLine As String
    Dim EmailBody As String
    
    FileNum = FreeFile()
    Open Filename For Input As #FileNum
    
    While Not EOF(FileNum)
        Line Input #FileNum, DataLine
        EmailBody = EmailBody & "<br>" & DataLine
    Wend
    
    FormatTemplateEmail = EmailBody
    
'Error Handling
catch:
    If err.Number <> 0 Then
        CatchErrorController "Tribbler.FormatTemplateEmail"
        Exit Function
    End If
End Function

'/////////////////////////SNOWFLAKE DB ARCHIVING FUNCTIONS //////////////////////
'CREATED: 01/04/2021
'@Author: nicholas.ackerman@albertsons.com
'PURPOSE: The purpose of these function is to ensure that data from the Snowflake tables are not lost.
'This function can be run on command and the data from the snowflake tools that are added can be archived
'This function should only be run by admin
'Save the data in a .csv archived file in the RT_MACROS Folder
'This file can be zipped to save space.
'Currently Archives TRIBBLES, CICS, and CLAIMS
'STORED IN K:\AA\SHARE\AuditTools\rtmacros\data\Archive
'Tool: = Claims or Tribble
Public Sub ArchiveSFDataBases(Optional Tool As String)
On Error GoTo catch
    Dim SaveFile As String
    
    'If Config.IsInsider = False Then
        'MsgBox "You do not have access to this function. Contact BA"
        'Exit Sub
    'End If
    
    'claims Archive
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    
    If Tool = "Claims" Or Tool = "" Then
        Sheets.Add.name = "CLAIMS"
        ImportSnowflakeTable "SELECT * FROM DW_PRD.TEMP_TABLES.PA_CLAIMS"
        
        'Remove files from Claims Folder
        Kill "K:\AA\SHARE\AuditTools\rtmacros\data\Archive\Claims\*.csv"
        
        Worksheets("CLAIMS").Activate
        SaveFile = "K:\AA\SHARE\AuditTools\rtmacros\data\Archive\Claims\CLAIMS" '& format(Now, "MM-DD-YYYY")
        ArchiveSF SaveFile
        Worksheets("CLAIMS").Delete
    End If
    
    If Tool = "Tribble" Or Tool = "" Then
        Sheets.Add.name = "TRIBBLES"
        ImportSnowflakeTable "SELECT * FROM DW_PRD.TEMP_TABLES.PA_TRIBBLES"
        Sheets.Add.name = "CICS"
        
        'Remove files from Tribbler Folder
        ImportSnowflakeTable "SELECT * FROM DW_PRD.TEMP_TABLES.PA_TRIBBLES_CICS"
        
        Kill "K:\AA\SHARE\AuditTools\rtmacros\data\Archive\Tribbler\*.csv*"
        'TRIBBLES
        Worksheets("TRIBBLES").Activate
        SaveFile = "K:\AA\SHARE\AuditTools\rtmacros\data\Archive\Tribbler\TRIBBLES" '& format(Now, "MM-DD-YYYY")
        ArchiveSF SaveFile
        Worksheets("TRIBBLES").Delete
        
        'TRIBBLE CICS
        Worksheets("CICS").Activate
        SaveFile = "K:\AA\SHARE\AuditTools\rtmacros\data\Archive\Tribbler\TRIBBLES_CICS" '& format(Now, "MM-DD-YYYY")
        ArchiveSF SaveFile
        Worksheets("CICS").Delete
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    MsgBox "Archiving Completed"
    Exit Sub
catch:
    If err.Number <> 0 Then
        CatchErrorController "Error Archiving: " & Tool
    End If
    MsgBox "Error Archiving"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
End Sub
'CREATED: 01/04/2021
'@Author: nicholas.ackerman@albertsons.com
'@Purpose: helperfunction for SFArchiving function
Public Sub ArchiveSF(ByVal SaveFile As String)
    Dim tempWB As Workbook
    ActiveSheet.Copy
    Set tempWB = ActiveWorkbook
    With tempWB
        .SaveAs Filename:=SaveFile, FileFormat:=xlCSV, CreateBackup:=False
        .Close
    End With
End Sub
'////////////////////////////TESTING NEW TOOL PACS OR CABS///////////////////////////////////////////
'Purpose: The purpose of this function is to take in two inputs of an offernumber and a
'divison. A query is then run and the result is either CABS or PACS for billing
'location. If True then in CABS else false is in PACS
'References: TestPacsOrCabs
'Author: Nicholas Ackerman
'Created: 01-19-21
Public Sub LookupPacsOrCabs()
    Static OfferNumber As String
    Static Division As String
    
    With New TwoInputboxForm
        .Caption = "PACS and WIMS Lookup"
        .Input1 = OfferNumber
        .Input2 = Division
        .show vbModal 'display the dialog
        If Not .IsCancelled Then 'how was it closed?
            OfferNumber = .Input1
            Division = .Input2
        End If
    End With
    
    If OfferNumber = vbNullString Or Division = vbNullString Then
        Exit Sub
    End If
    
    Dim ans As Boolean
    ans = CABS.TestPacsOrCabs(OfferNumber, Division)
    If ans = True Then
        MsgBox "OfferNumber: " & OfferNumber & " is located in the CABS system"
    Else
        MsgBox "OfferNumber: " & OfferNumber & " is located in the PACS system"
    End If
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 04/08/2021
'Purpose:Test to see if I can remove many different types of special characters that might break a query
'Currently this function removes "#'%^*&"
'References:
Public Function RemoveSpecialChars(Str As String) As String
    Dim xChars As String
    Dim index As Long
    xChars = "#'%^*&"
    For index = 1 To Len(xChars)
        Str = Replace$(Str, Mid$(xChars, index, 1), "")
    Next
    
    RemoveSpecialChars = Str
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 06/01/2021
'Purpose:WORKS WITH A JASON FILE, True = Admin, False = None-Admin
'References:JsonParse, ReadTextFile, Scripting Dictionary, Environ
'This must take a list of admin variables from a JSON file.
Public Function IsUserAdmin(ByRef JSONFile As String) As Boolean
    Dim Data As Scripting.Dictionary
    Set Data = JsonParse(ReadTextFile(ENV.use(JSONFile)))
    
    'Admin Users Function to populate additional functions
    If Contains(Data("admin"), UCase(Environ("Username"))) = True Then
        IsUserAdmin = True
    Else
        IsUserAdmin = False
    End If
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 06/23/2021
'Purpose:The purpose ofthis function is to generate the Tomorrow Is Emails. It pulls the claims over 10,000 for the current day and places them into a email.
Public Sub TomorrowIs()
    'Function to automate the generation of Lindsay's Tomorrow Is Emails
    Dim rs As ADODB.Recordset
    Dim finalEmail As String
    Dim EmailBody As String
    Dim Data As Scripting.Dictionary
    Dim Todays_Date As String
    
    LogCode "TomorrowIs"            'Log Reference to Tool
    
    Dim USER_DATE As String
    Todays_Date = Month(Now()) & "/" & Day(Now()) & "/" & YEAR(Now())
    USER_DATE = InputBox("PLEASE ENTER THE DATE YOU WISH TO RUN THE REPORT FOR [MM/DD/YYYY]: ", "INPUT_DATE", Todays_Date)
    If IsDate(USER_DATE) = False Then
        MsgBox "Please Enter a Valid Date MM/DD/YYYY"
        Exit Sub
    End If
    
    Todays_Date = YEAR(USER_DATE) & "-" & Month(USER_DATE) & "-" & Day(USER_DATE)
    
    Set Data = JsonParse(ReadTextFile(ENV.use("CONFIG_CLAIMS")))
    Set rs = QuerySnowFlake("SELECT * FROM DW_PRD.TEMP_TABLES.PA_CLAIMS_CLONE WHERE DATE_CREATED = '" & Todays_Date & "' AND AMOUNT >= 10000.00 ORDER BY CLAIMED DESC")
    
    If IsRecordsetEmpty(rs) Then
        MsgBox "No Claims were made over 10,000 today. Sorry"
        Exit Sub
    End If
    
    finalEmail = "<html><body lang=EN-US link='#0563C1' vlink='#954F72' style='word-wrap:break-word'><div class=WordSection1>"
    
    Do While Not rs.EOF
        Dim tempString As String
        EmailBody = FormatTemplateEmail("K:\AA\SHARE\AuditTools\rtmacros\data\EmailTemplates\Tomorrow_Is.txt")
        tempString = "<b>" & Data(UCase(rs.Fields("CLAIMED").Value)) & "</b> found an <b>" & rs.Fields("REASON").Value & "</b> from <b>" & rs.Fields("VENDOR_NAME").Value & "</b> DIV:<b>" & rs.Fields("DIV").Value & "</b>"
        
        'replace values
        EmailBody = Replace(EmailBody, "(&USER)", tempString)
        EmailBody = Replace(EmailBody, "(&AMOUNT)", "$" & format(CStr(rs.Fields("AMOUNT").Value), "#,###"))
        finalEmail = finalEmail & EmailBody
        rs.MoveNext
    Loop
    
    finalEmail = finalEmail & "</div></body></html>"

    'BasicOutlookEmailCC Replace(ArrayToString(Data("EmailList")), ",", ";"), Replace(ArrayToString(Data("CCEmailList")), ",", ";"), "Tomorrow Is", finalEmail
    BasicOutlookEmailCC JSONArrayUsernameToEmailString(ENV.use("CONFIG_CLAIMS"), "EmailUsers"), JSONArrayUsernameToEmailString(ENV.use("CONFIG_CLAIMS"), "CCEmailUsers"), "Tomorrow Is", finalEmail
    
End Sub
Public Function IsLoaded(formName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.name = formName Then
            IsLoaded = True
            Exit Function
        End If
    Next frm
    IsLoaded = False
End Function
Public Function ColumnNumber2Letter(ByRef ColumnNumber As Long) As String
'PURPOSE: Convert a given number into it's corresponding Letter Reference
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    Dim ColumnLetter As String
    
    'Convert To Column Letter
    ColumnLetter = split(Cells(1, ColumnNumber).Address, "$")(1)
      
    'Display Result
    ColumnNumber2Letter = ColumnLetter
End Function
Public Function ColumnRangeToText(Rng As String, delimiter As String) As String
    Dim c As Variant
    For Each c In ActiveSheet.Range(Rng).Cells
        ColumnRangeToText = ColumnRangeToText & c.Value & delimiter
    Next
    
    ColumnRangeToText = Left(ColumnRangeToText, Len(ColumnRangeToText) - Len(delimiter))

End Function
Public Function PullEmailFromUsername(ByRef Username As String) As String
    Dim DataEmails As Scripting.Dictionary
    Dim Email As Scripting.Dictionary
    Set DataEmails = JsonParse(ReadTextFile(ENV.use("CONFIG_EMAILS")))
       
    If DataEmails.Exists(UCase(Username)) Then
        Set Email = DataEmails(UCase(Username))
        PullEmailFromUsername = Email("EMAIL")
    End If
End Function
Public Function JSONArrayUsernameToEmailString(ByRef JSONFilePath As String, ByRef UsernameArray As String) As String
    Dim Data As Scripting.Dictionary
    Set Data = JsonParse(ReadTextFile(JSONFilePath))
    Dim UsernameList() As Variant
    Dim Username As Variant
    
    UsernameList = Data(UsernameArray)
    For Each Username In UsernameList
        JSONArrayUsernameToEmailString = PullEmailFromUsername(CStr(Username)) & ";" & JSONArrayUsernameToEmailString
    Next Username
    
    JSONArrayUsernameToEmailString = Left(JSONArrayUsernameToEmailString, Len(JSONArrayUsernameToEmailString) - 1)
    
End Function
'The purpose of this function is to add a user and email to the email JSON file.
'12/07/21
Public Sub AddUserToEmailListConfig(Username As String, Email As String)
    Dim FILEPATH As String
    FILEPATH = "K:\AA\SHARE\AuditTools\rtmacros\data\Users_Emails.json"
    
    Dim Settings As Scripting.Dictionary
    Dim EmailSettings As New Scripting.Dictionary
    Set Settings = JsonParse(ReadTextFile(FILEPATH))
    
    'CHECK TO SEE IF USER EXISTS IN EMAIL SETTINGS
    If Not Settings.Exists(Username) Then
        EmailSettings.Add "EMAIL", Email
        Set Settings(Username) = EmailSettings
        WriteToTextFile FILEPATH, JsonStringify(Settings)
        
    End If
    
End Sub
'The purpose of this function is to add a user to TRIBBLE CONFIGURATION FILE.
'12/07/21
'DETEMRINE IF THE USER IS AN ADMIN OR A BOUNCER AS WELL.
'This will not work with a bouncer lead or tribbler positions
Public Sub AddUserToTribbleConfig(Username As String, name As String)
    Dim FILEPATH As String
    FILEPATH = "K:\AA\SHARE\AuditTools\rtmacros\sql\Tribbler\ConfigTribble.json"
    
    Dim Settings As Scripting.Dictionary
    Dim DataArray() As Variant
    Set Settings = JsonParse(ReadTextFile(FILEPATH))

    'Add Admin to Tribbler
    If MsgBox("Will this user be an Admin for the TRIBBLER tool?", vbYesNo) = vbYes Then
        DataArray = Settings("admin")
        ArrayPush DataArray, Username
        Settings("admin") = DataArray()
        WriteToTextFile FILEPATH, JsonStringify(Settings)
    End If
    
    'Add user to Tribbler
    DataArray = Settings("users")
    ArrayPush DataArray, Username
    Settings("users") = DataArray()
    WriteToTextFile FILEPATH, JsonStringify(Settings)
    
    'Add user to Bouncers
    If MsgBox("Will this user be a Bouncer for the TRIBBLER tool?", vbYesNo) = vbYes Then
        DataArray = Settings("bouncers")
        ArrayPush DataArray, Username
        Settings("bouncers") = DataArray()
        WriteToTextFile FILEPATH, JsonStringify(Settings)
    End If

End Sub
'The purpose of this function is to add a user to CLAIMS CONFIGURATION FILE.
'12/07/21
Public Sub AddUserToClaimsConfig(Username As String)
    Dim FILEPATH As String
    FILEPATH = "K:\AA\SHARE\AuditTools\rtmacros\sql\Claims\ConfigClaims.json"
    
    Dim Settings As Scripting.Dictionary
    Dim DataArray() As Variant
    Set Settings = JsonParse(ReadTextFile(FILEPATH))
    
    'Add Admin to Claims tool
    If MsgBox("Will this user be an Admin for the CLAIMS tool?", vbYesNo) = vbYes Then
        DataArray = Settings("admin")
        ArrayPush DataArray, Username
        Settings("admin") = DataArray()
        WriteToTextFile FILEPATH, JsonStringify(Settings)
    End If
    
    'Add user to Claims tool
    DataArray = Settings("users")
    ArrayPush DataArray, Username
    Settings("users") = DataArray()
    WriteToTextFile FILEPATH, JsonStringify(Settings)

End Sub
Public Sub AddUserToBCConfigPreAudit(Username As String)
    Dim FILEPATH As String
    FILEPATH = "K:\AA\SHARE\AuditTools\rtmacros\sql\BillingCorrespondence\ConfigBC.json"
    
    Dim Settings As Scripting.Dictionary
    Dim DataArray() As Variant
    Set Settings = JsonParse(ReadTextFile(FILEPATH))

    'Add Admin to BC CONFIG FILE FOR
    DataArray = Settings("EmailUsers")
    ArrayPush DataArray, Username
    Settings("EmailUsers") = DataArray()
    WriteToTextFile FILEPATH, JsonStringify(Settings)

End Sub
'The purpose of this function is to add a user to CLAIMS CONFIGURATION FILE.
'12/07/21
Public Sub AddUserToUsersConfig(Username As String)
    Dim FILEPATH As String
    FILEPATH = "K:\AA\SHARE\AuditTools\rtmacros\users.json"
    
    Dim Settings As Scripting.Dictionary
    Dim VerisonNum As New Scripting.Dictionary
    Set Settings = JsonParse(ReadTextFile(FILEPATH))
    
    VerisonNum.Add "version", CStr(VersionControl.VersionNumber)
    Set Settings(Username) = VerisonNum
    WriteToTextFile FILEPATH, JsonStringify(Settings)
    
   
End Sub
'Created: 12/14/2021
'Nick Ackerman
'Purpose is to check to see if the user has access to RT_MACROS
Public Function CheckIfUserInRT_Macros(Username As String) As Boolean
    Dim FILEPATH As String
    Dim Settings As Scripting.Dictionary
    FILEPATH = "K:\AA\SHARE\AuditTools\rtmacros\users.json"
    
    Set Settings = JsonParse(ReadTextFile(FILEPATH))
    
    If Settings.Exists(Username) = True Then
        CheckIfUserInRT_Macros = True
    Else
        CheckIfUserInRT_Macros = False
    End If
    
End Function
'Created: 12/14/2021
'Nick Ackerman
'Purpose is to add a new user to the Tribbler, Claims, and Email List
Public Sub AddNewUserToPreAuditTools()
    Dim Username As String
    Dim FirstName As String
    Dim EmailAddress As String
    Username = UCase(InputBox("Please Enter New Users Username"))
    
    If CheckIfUserInRT_Macros(Username) = False Then
        FirstName = UCase(InputBox("Please Enter New Users FirstName"))
        EmailAddress = InputBox("Please Enter New Users Email Address")
        
        'Add users to files
        AddUserToUsersConfig Username
        AddUserToBCConfigPreAudit Username
        AddUserToClaimsConfig Username
        AddUserToTribbleConfig Username, FirstName
        AddUserToEmailListConfig UCase(Username), EmailAddress
        MsgBox "User added to USERS, Claims, Tribbler, and  Email List."
    Else
        MsgBox "User already has access to RT_MACROS"
    End If
    
End Sub

'The purpose of this function is to add a user to BILLING CORRESPONDENCE CONFIGURATION FILE.
'12/07/21
Public Sub AddUserToBCConfig()
    Dim FILEPATH, Username, fullname As String
    FILEPATH = "K:\AA\SHARE\AuditTools\rtmacros\sql\Claims\ConfigClaims.json"
    Username = "JDOEE04"
    fullname = "JOHN DOE"
    
    Dim Settings As Scripting.Dictionary
    Dim DataArray() As Variant
    Set Settings = JsonParse(ReadTextFile(FILEPATH))
    
    'Add user to general identifer
    Set Settings(Username) = Username
    WriteToTextFile FILEPATH, JsonStringify(Settings)


    'Add Admin to Tribbler
    DataArray = Settings("EmailUsers")
    ArrayPush DataArray, Username
    Settings("EmailUsers") = DataArray()
    WriteToTextFile FILEPATH, JsonStringify(Settings)
    
    
    'Add user to Tribbler
    DataArray = Settings("CCEmailUsers")
    ArrayPush DataArray, Username
    Settings("CCEmailUsers") = DataArray()
    WriteToTextFile FILEPATH, JsonStringify(Settings)
    
    'Add users full name to Biller List
    DataArray = Settings("BillerList")
    ArrayPush DataArray, fullname
    Settings("BillerList") = DataArray()
    WriteToTextFile FILEPATH, JsonStringify(Settings)
    
    'Add users username to Biller List USRM
    DataArray = Settings("BillerListUSRM")
    ArrayPush DataArray, Username
    Settings("BillerListUSRM") = DataArray()
    WriteToTextFile FILEPATH, JsonStringify(Settings)
    
    'Add users username to ACCOUNTANT List USRM
    DataArray = Settings("AccountantListUSRNM")
    ArrayPush DataArray, Username
    Settings("AccountantListUSRNM") = DataArray()
    WriteToTextFile FILEPATH, JsonStringify(Settings)
    
    'Add users username to admin developer for Billing tool
    DataArray = Settings("ADMIN_DEV")
    ArrayPush DataArray, Username
    Settings("ADMIN_DEV") = DataArray()
    WriteToTextFile FILEPATH, JsonStringify(Settings)

End Sub
'produce a function to undo anything that is uploaded if an error occurs
'This is selenium based function
'Future testing will be with metric.
Public Sub seleniumtutorial()
    Dim bot As New WebDriver
    With bot
        .start "edge", "https://mydt.demandtec.com/dteai/login?TAM_OP=login&ERROR_CODE=0x00000000&METHOD=GET&URL=%2F&HOSTNAME=mydt.demandtec.com&AUTHNLEVEL=&FAILREASON=&dtuser=Unknown"
        .Get "/"
        .FindElementByName("username").Sendkeys ("NACKE08")
        .FindElementByName("password").Sendkeys ("ERRor")
        .FindElementByCss(".btn").Click
        .FindElementById("sub-li-dealmanagement").Click
        .FindElementById("AutoComplete0_TextField").Sendkeys ("4491772")
        .FindElementById("btnFSecFind").Click
        .FindElementByLinkText("View").Click
        Debug.Print ("Test")
        .ExecuteScript ("window.print()")
        .Wait 30000
    End With
    
    'bot.Quit
    'quit method for closing browser instance.
End Sub
'Updated 01/05/2022
'Purpose The purpose is to pull in all of the offer/cic information for a group offers inserted into the tool
'Author: Nicholas Ackerman
Public Sub PullOfferDetails(OfferNumbers As String)
    Dim SQL As String
    Dim row As Integer
    Dim DataSheetName As String
    Dim NewSheetName As String
    Dim SelectionAddress As String
    
    'Add a new Sheet
    Sheets.Add
    
    'Build SQL String
    SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\OFFER_DETAILS_TOOL_V3.txt")
    SQL = Replace(SQL, "(&OFFER_NBRS)", OfferNumbers)
    
    'Import into Spreadsheet
    ImportSnowflakeTable SQL

End Sub

'Updated 01/05/2022
'Purpose 'This function can be used in conjunction with the OFFER DETAILS TOOL. This will provide the div and offer number of the overlapping offer/offers
'Author: Nicholas Ackerman
Public Sub PullOverlappingOfferDetails()
    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim row As Integer
    Dim DataSheetName As String
    
    If IsNumeric(Selection.Value) = False Then
        MsgBox "Please Select a Valid OfferNumber"
        Exit Sub
    End If
    
    'Build SQL String
    SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\OVERLAPPING_OFFERS_V1.txt")
    SQL = Replace(SQL, "(&OFFER_NBRS)", Selection.Value)

    Set rs = QuerySnowFlake(SQL)
    
    If IsRecordsetEmpty(rs) Then
        MsgBox ("No OVERLAPPING FOUND")
        Dialog.Hide
        Exit Sub
    End If
    'BUILD A MESSAGE BOX WITH DATA
    Dim message As String
    message = "OVERALPPING OFFERS" & vbNewLine & "DIV" & vbTab & "--" & vbTab & "OFFER" & vbNewLine
        
    Do While Not rs.EOF
        message = message & rs.Fields("DIV").Value & vbTab & "--" & vbTab & rs.Fields("OL_OFFER").Value & vbNewLine
        rs.MoveNext
    Loop
    
    MsgBox message

End Sub
'Updated 01/05/2022
'Purpose 'This function is used to pull all of the offer numbers for the S2SQ and ScanQ Queries. This is used in the SingleListFormViewOffers Tool
'Author: Nicholas Ackerman
Public Function PullOfferNumbersFromS2S_Scan(Letter As String, SDATE As String, EDATE As String) As String
    'Run the Query to pull the offer number for S2SQ and ScanQ
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim OfferString As String
    SQL = ReadTextFile("K:\AA\SHARE\AuditTools\rtmacros\sql\SCAN_S2S_Q_OFFERS_ONLY.sql")
    SQL = Replace(SQL, "(&LETTER)", Letter)
    SQL = Replace(SQL, "(&SDATE)", SDATE)
    SQL = Replace(SQL, "(&EDATE)", EDATE)
    
    Set rs = QuerySnowFlake(SQL)
    
    Do While Not rs.EOF
        OfferString = rs.Fields("OFFER_NUMBER").Value & "," & OfferString
        rs.MoveNext
    Loop
    
    PullOfferNumbersFromS2S_Scan = Left(OfferString, Len(OfferString) - 1)
    
End Function
'credits StackOverflow User :)
'Created: 10/27/2020
'Purpose:'Used to clear spaces before and after a string
Public Function ClearExtraSpaces(InputString As String) As String

    Dim RE As RegExp
    Set RE = New RegExp
    With RE
        .Global = True
        .MultiLine = True
        .pattern = "^\s*(\S.*\S)\s*"
        InputString = .Replace(InputString, "$1")
    End With
    ClearExtraSpaces = InputString

End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 12/10/2020
'Purpose:This is called when a picture needs to be popualted on a surprise congrats form
'References: K:\AA\SHARE\AuditTools\rtmacros\images\SURPRISE
Public Sub SurpriseFormRandomPicture()
On Error GoTo catch
    Const Path As String = "K:\AA\SHARE\AuditTools\rtmacros\images\SURPRISE\"
    Dim Filename As String
    Dim numOfFiles As Integer
    Dim inputFiles As Variant

    Dim fso As New Scripting.FileSystemObject
    Dim randomNum As Integer

    Randomize
    randomNum = RandomNumber(1, CInt(fso.GetFolder(Path).Files.Count))

    Dim pictureFile As file
    Dim Count As Integer
    Count = 1
    For Each pictureFile In fso.GetFolder(Path).Files
        If Count = randomNum Then
            Filename = pictureFile.name
        End If
        Count = Count + 1
    Next pictureFile

    CongratsForm.Picture = LoadPicture(Path & Filename)
    CongratsForm.show
    Exit Sub
catch:
    CongratsForm.Picture = LoadPicture("K:\AA\SHARE\AuditTools\rtmacros\images\SURPRISE\HappyGolden.jpg")
    CongratsForm.show
End Sub
'Created: 12/17/2021
'Author: Nicholas Ackerman
'Purpose: The purpsoe of this function is to determine if an offer number resides in CABS or in PACS. The return value is a string from the CABS DB that
'determines if the offer is CABS or PACS.
'This is used to determine which query to run for the Claims tool
Public Function IsCABSorPACS(OfferNum As String) As String
    Dim SQL As String
    On Error GoTo catch
    Dim rs As ADODB.Recordset
    SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\CABS_OR_PACS_V2.sql")
    SQL = Replace(SQL, "(&OFFER_NUMBER)", OfferNum)

    Set rs = QueryCABS(SQL)
    
    Do While Not rs.EOF
        If rs.Fields("STATUS").Value = "PACS" Then
            IsCABSorPACS = "CABS"
        Else
            IsCABSorPACS = "PACS"
        End If
        rs.MoveNext
    Loop
    
'Error Handling
    Exit Function
catch:
    Console.error err.Description, "IsCABSorPACS"
    DisplayErrorMessage "IsCABSorPACS" & vbCr & vbCr & err.Description
    
End Function
'Created: 01/27/2021
'Nicholas Ackerman
'Purpose: Is to pull all of the private and public Subs/Function in RT_Macros into a log file that is stored at K:\AA\SHARE\AuditTools\rtmacros\data\Archive
' the file is called FUNCTIONS_SUBS.LOG.
    'types
    '1 = Module
    '2 = class module
    '3 = forms
    '100 = workbook
Public Sub CreateListOfPUBLICFunctionsAndSubs()
    Dim Path As String
    Dim FileNumber As Integer
    
    Path = "K:\AA\SHARE\AuditTools\rtmacros\data\Archive\PUBLIC_FUNCTIONS_SUBS.LOG"
    FileNumber = FreeFile
    Open Path For Output As FileNumber
    
    Dim c As VBComponent
    For Each c In ThisWorkbook.VBProject.VBComponents
        If c.Type <> "3" And c.Type <> "100" Then
            Print #FileNumber, c.name
            Print #FileNumber, "Members:"
            Dim ln As Long
            For ln = 1 To c.CodeModule.CountOfLines
                Dim lineOfCode As String
                lineOfCode = c.CodeModule.lines(ln, 1)
                
                If Left(lineOfCode, 4) <> "End " And (InStr(lineOfCode, "Public Sub") > 0 Or InStr(lineOfCode, "Public Function") > 0) Then
                    Print #FileNumber, vbTab & split(lineOfCode, "(")(0)
                End If
            Next
        End If
        Print #FileNumber, vbNewLine
    Next
    Close FileNumber
End Sub
'Created: 01/27/2021
'Nicholas Ackerman
'Purpose:
Public Function BouncerReportVersion2()
    On Error Resume Next
    Application.EnableEvents = True
    Workbooks.Open "C:\Users\nacke08\OneDrive - Safeway, Inc\Documents\Projects\BOUNCER_PROJECT\TEMPLATE_V1_NICK.xlsm", , True, , , , True
End Function
