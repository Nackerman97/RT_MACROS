Attribute VB_Name = "Claims_V2"
Option Explicit

'//------------------------------------------------------------------------------------------------------------------//
'//------------------------------------------------------------------------------------------------------------------//
'Version 2 of Claims Tool below this line section
'TODO: Work on the Error Handling. '01/25/2022
'//------------------------------------------------------------------------------------------------------------------//
'//------------------------------------------------------------------------------------------------------------------//
Public Enum claimsDB2Information
    DIV = 0
    Vendor = 1
    Log = 2
    VENDOR_NAME = 3
    ALLOW_TYPE = 4
    PAY_TYPE = 5
    PERFORM_1 = 6
    PERFORM_2 = 7
    FAC_NBR = 8
    SF_NBR = 9
    CCOA_ACCT_NBR = 10
    FAC_NBR2 = 11
    SF_NBR2 = 12
    CCOA_ACCT_NBR3 = 13
    MERCHANDISER = 14
    VDC_NBR = 15
    VDC_DTE = 16
    CURRENT_DATE = 17
    Tribble_ID = 18
    
End Enum
Public Sub LocateClaims()
    
    ValidateClaimsLookUpField_V2
    
    If ClaimsInputV2.OfferNumTextBox.Value <> "" Then
        QueryForClaims_V2 ClaimsInputV2.OfferNumTextBox.Value
    End If
    
End Sub
Public Sub LocateClaimsInSF_V2(SQL As String)
On Error GoTo catch
    Dim rs As ADODB.Recordset
    Dim row As Integer
    
    With ClaimsInputV2.ClaimsListBox
        .AddItem
        .List(0, 0) = "DIV"
        If InStr(ClaimsInputV2.OfferNumTextBox.Value, "D") > 0 Then
            .List(0, 1) = "DEBIT_MEMO"
        Else
            .List(0, 1) = "OFFER_NUMBER"
        End If
        .List(0, 2) = "AMOUNT"
        .List(0, 3) = "VDC_DTE"
        .List(0, 4) = "VDC_NBR"
        .List(0, 5) = "BILLER_ERROR"
        .List(0, 6) = "TRIBBLE"
        .List(0, 7) = "PERIOD"
        .List(0, 8) = "YEAR"
        .List(0, 9) = "REASON"
    End With
    
    Set rs = QuerySnowFlake(SQL)
    row = 1
    Do While Not rs.EOF
        With ClaimsInputV2.ClaimsListBox
            .AddItem
            .List(row, 0) = rs.Fields("DIV").Value
            If InStr(ClaimsInputV2.OfferNumTextBox.Value, "D") > 0 Then
                .List(row, 1) = rs.Fields("VDC_NBR").Value
            Else
                .List(row, 1) = rs.Fields("OFFER_NUMBER").Value
            End If
            .List(row, 2) = rs.Fields("AMOUNT").Value
            .List(row, 3) = rs.Fields("VDC_DTE").Value
            .List(row, 4) = rs.Fields("VDC_NBR").Value
            .List(row, 5) = rs.Fields("BILLER_ERROR").Value
            .List(row, 6) = rs.Fields("TRIBBLE").Value
            .List(row, 7) = rs.Fields("PERIOD").Value
            .List(row, 8) = rs.Fields("YEAR").Value
            .List(row, 9) = rs.Fields("REASON").Value & "---" & rs.Fields("CLAIM_ID").Value '& "---" & rs.Fields("TRIBBLE_ID").value
        End With
        row = row + 1
        rs.MoveNext
    Loop
    
'Error Handling
    Exit Sub
catch:
    Console.error err.Description, "Claims.LocateClaimsInSF_V2"
    DisplayErrorMessage "Claims.LocateClaimsInSF_V2" & vbCr & vbCr & err.Description
        
End Sub
Public Sub RemoveClaimsFromSF_V2()
On Error GoTo catch
    Dim DT As String
    Dim Amount As String
    Dim OfferNumber As String
    Dim DIV As String
    Dim SQL As String
    
    DT = split(ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 9), "---")(1)
    Amount = ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 2)
    OfferNumber = ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 1)
    DIV = ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 0)
    
    '2021-06-17 09:31:53.000 Formatting Example
    DT = CStr(YEAR(DT)) & "-" & CStr(Month(DT)) & "-" & CStr(Day(DT)) & " " & CStr(Hour(DT) + 1) & ":" & CStr(Minute(DT)) & ":" & CStr(Second(DT)) & ".000"

    If InStr(UCase(ClaimsInputV2.OfferNumTextBox.Value), "D") > 0 Then
        SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_Delete_DM_SF_TS.txt")
        SQL = ClaimsSQLStringReplace(SQL, DIV, , , , , Amount, , , , , , , , , , , , , , , OfferNumber, , , , , , DT, "")
    Else
        SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_Delete_SF_TS.txt")
        SQL = ClaimsSQLStringReplace(SQL, DIV, , , OfferNumber, , Amount, , , , , , , , , , , , , , , , , , , , , DT, "")
    End If
    
    If MsgBox("Are you sure you want to Delete this Claim? It cannot be recovered. ", vbYesNo) = vbYes Then
        QuerySnowFlake (SQL)
        MsgBox "The Claim has been deleted"
    End If
    
'Error Handling
    Exit Sub
catch:
    Console.error err.Description, "Claims.RemoveClaimsFromSF_V2"
    DisplayErrorMessage "Claims.RemoveClaimsFromSF_V2" & vbCr & vbCr & err.Description
    
End Sub
Public Sub InsertClaimIntoSF_V2()
On Error GoTo catch
    Dim Data As Scripting.Dictionary
    Set Data = JsonParse(ReadTextFile(ENV.use("CONFIG_CLAIMS")))
    Dim SplitString() As String
    
    'Convert Header
    ValidateClaimsLookUpField_V2
    
    'Validation
    If ClaimsSelectionIndex_V2 = 0 Then
        If MsgBox("Would you like to proceed with a manual input?", vbYesNo) = vbYes Then
            MaunalInputFunction_V2
            Exit Sub
        Else
            Exit Sub
        End If
    ElseIf ValidateClaimsFieldsNotEmpty_V2 = False Then
        MsgBox "One of the Form Fields is Empty"
        Exit Sub
    End If
    
    'Check if the Claim Exists
     If CheckIfClaimExists_V2(ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 0), ClaimsInputV2.OfferNumTextBox.Value, ClaimsInputV2.AmountTextBox.Value) = False Then
        If InsertDataIntoSnowflake(MakeClaimQuery_V2) = 0 Then
            
        
            'Claim submitted correctly
            MsgBox "The Claim below has been submitted:" & vbCr & _
            ClaimsInputV2.OfferNumTextBox.Value & vbTab & "Amount: $" & ClaimsInputV2.AmountTextBox.Value
            
            'check if archiving is needed
            'CheckIfArchiveIsReady_V2
            
            'Output Surprise Picture
            SplitString = split(CStr(ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 9)), "--")
            ClaimsTolerancesPopUpPicture_V2 CStr(SplitString(5)), Data
            
        End If
     End If
    
    ClearClaimFormBoxes_V2
    
'Error Handling
    Exit Sub
catch:
    Console.error err.Description, "Claims.InsertClaimIntoSF_V2"
    DisplayErrorMessage "Claims.InsertClaimIntoSF_V2" & vbCr & vbCr & err.Description
    
End Sub
Public Sub PopulateClaimsFormFields_V2(UpdateIdentification As Integer)
    On Error GoTo catch
    Dim index As Integer
    Dim RowData() As String
    With ClaimsInputV2
        For index = 0 To .ClaimsListBox.ListCount - 1
            If .ClaimsListBox.Selected(index) = True And index <> 0 Then
    
                .AmountTextBox.Value = .ClaimsListBox.List(index, 2)
                .PeriodComboBox.Value = .ClaimsListBox.List(index, 7)
                .YearComboBox.Value = .ClaimsListBox.List(index, 8)
                
                'Reasons
                If UpdateIdentification = 1 Then
                'If Contains(.ClaimsListBox.List(index, 9), "---") = True Then
                    .ReasonComboBox.Value = split(.ClaimsListBox.List(index, 9), "---")(0)
                    .IDTextBox.Value = split(.ClaimsListBox.List(index, 9), "---")(1)
                    '.TribbleTextBox.value = split(.ClaimsListBox.List(index, 9), "---")(2)
                Else
                    .ReasonComboBox.Value = ""
                End If
    
                'Conversion for Biller Error TextBox
                If .ClaimsListBox.List(index, 5) = "Yes" Or _
                .ClaimsListBox.List(index, 5) = "TRUE" Or _
                .ClaimsListBox.List(index, 5) = -1 Then
                    .BillerErrorComboBox.Value = "True"
                Else
                    .BillerErrorComboBox.Value = "False"
                End If
                'Conversion for Tribble TextBox
                If .ClaimsListBox.List(index, 6) = "Yes" Or _
                .ClaimsListBox.List(index, 6) = "TRUE" Or _
                .ClaimsListBox.List(index, 6) = -1 Then
                    .TribbleComboBox.Value = "True"
                Else
                    .TribbleComboBox.Value = "False"
                End If
                
            End If
        Next index
    End With
    
'Error Handling
    Exit Sub
catch:
    Console.error err.Description, "Claims.PopulateClaimsFormFields_V2"
    DisplayErrorMessage "Claims.PopulateClaimsFormFields_V2" & vbCr & vbCr & err.Description
    
End Sub

Public Function ClaimsSearchSQL_V2() As String
    
    If IdentifyClaimsType_V2 = "DM" Then
        ClaimsSearchSQL_V2 = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Search_DebitMemo.txt")
        ClaimsSearchSQL_V2 = Replace(ClaimsSearchSQL_V2, "(&OFFER_NUMBER)", ClaimsInputV2.OfferNumTextBox.Value)
    Else
        ClaimsSearchSQL_V2 = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Search_OfferNum.txt")
        ClaimsSearchSQL_V2 = Replace(ClaimsSearchSQL_V2, "(&OFFER_NUMBER)", ClaimsInputV2.OfferNumTextBox.Value)
    End If
    
End Function
'Purpose: This sub compiles the offernumber in Db2 to find the claim
'References: FillClaimFormBoxes,claimsDB2Data, IsCABSorPACS
'Check if the Query shoulf be checked with DB2 or CABS
Public Sub QueryForClaims_V2(OfferNum As String)
    
    If InStr(OfferNum, "D") > 0 Or IsCABSorPACS(OfferNum) = "PACS" Then
        FillClaimFormBoxes_V2 claimsDB2Data_V2(OfferNum)
    Else
        FillCABSData OfferNum
    End If

End Sub
'Purpose: This function runs the code to search for the invoices in the claims sheet
'This function accesses DB2
'References:DB2.RunQuery
Public Function claimsDB2Data_V2(OfferNumber As String) As Variant
    Dim SQL As String
    
    If InStr(OfferNumber, "D") > 0 Then
        'To pull row information for DB2 Debit Memo
        SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_DebitMemo.txt")
        SQL = Replace(SQL, "(&DEBIT_MEMO_NUMBER)", OfferNumber)
    Else
        'To pull row information for DB2
        SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims.txt")
        SQL = Replace(SQL, "(&OFFER_NUMBER)", OfferNumber)
    End If
    
    claimsDB2Data_V2 = DB2.RunQuery(SQL)
        
End Function
'Purpose:This sub filles in the Listbox in the main form.
'References:ClaimsInput
'DB2 queries for DebitMemo and Offer Number
Public Sub FillClaimFormBoxes_V2(Data As Variant)
    On Error GoTo catch
    
    ClaimsInputV2.ClaimsListBox.Clear
    
    Dim size As Double
    size = ClaimsInputV2.ClaimsListBox.Width / 5
    ClaimsInputV2.ClaimsListBox.ColumnWidths = CStr(size) & "," & CStr(size) & "," & CStr(size) & "," & CStr(size) & "," & CStr(size) & "1,1,1,1,1"

    Dim row As Integer
    For row = 0 To arrayLength(Data) - 1
        With ClaimsInputV2.ClaimsListBox
            .AddItem
            .List(row, 0) = Data(row, 0)  'DIV
            .List(row, 1) = Data(row, 3)  'OFFER NUM
            .List(row, 2) = Data(row, 6)  'AMOUNT
            .List(row, 3) = Data(row, 24) 'VDC_DTE
            .List(row, 4) = Data(row, 23) 'VDC_NBR
            .List(row, 5) = Data(row, 8)    'biller_error
            .List(row, 6) = Data(row, 9)    'tribble
            .List(row, 7) = Data(row, 21) ' year
            .List(row, 8) = Data(row, 22) 'period
            .List(row, 9) = CStr(Data(row, 0)) & "--" & CStr(Data(row, 1)) & "--" & CStr(Data(row, 2)) & "--" & CStr(Data(row, 4)) & "--" & CStr(Data(row, 10)) & "--" & CStr(Data(row, 11)) & "--" & _
                            CStr(Data(row, 12)) & "--" & CStr(Data(row, 13)) & "--" & CStr(Data(row, 14)) & "--" & CStr(Data(row, 15)) & "--" & CStr(Data(row, 16)) & "--" & CStr(Data(row, 17)) & "--" & _
                            CStr(Data(row, 18)) & "--" & CStr(Data(row, 19)) & "--" & CStr(Data(row, 20)) & "--" & CStr(Data(row, 23)) & "--" & CStr(Data(row, 24)) & "--" & CStr(Data(row, 25))
                            
        End With
    Next row
    
    Exit Sub
catch:
    Dim ErrorMessage As String
    ErrorMessage = err.Description

    Application.ScreenUpdating = True
    Console.error ErrorMessage, "Claims.FillClaimFormBoxes"
    
    DisplayErrorMessage ErrorMessage
End Sub
'Created 12/17/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose: Is to fill the claims box with the information from the CABS DB.
Private Sub FillCABSData(OfferNum As String)
On Error GoTo catch
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim row As Integer
    SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_Search_OfferNum_CABS.txt")
    SQL = Replace(SQL, "(&OFFER_NUMBER)", OfferNum)
    
    Set rs = QueryCABS(SQL)
    
    ClaimsInput.ClaimsListBox.Clear
    
    Dim size As Double
    size = ClaimsInputV2.ClaimsListBox.Width / 5
    ClaimsInputV2.ClaimsListBox.ColumnWidths = CStr(size) & "," & CStr(size) & "," & CStr(size) & "," & CStr(size) & "," & CStr(size) & "1,1,1,1,1"
    
    With ClaimsInputV2.ClaimsListBox
        .AddItem
        .List(0, 0) = "DIV"
        .List(0, 1) = "OFFER_NUMBER"
        .List(0, 2) = "AMOUNT"
        .List(0, 3) = "VDC_DTE"
        .List(0, 4) = "VDC_NBR"
        .List(0, 5) = "BILL_ERROR"
        .List(0, 6) = "TRIBBLE"
        .List(0, 7) = "PERIOD"
        .List(0, 8) = "YEAR"
        .List(0, 9) = "REASON"
    End With
    
    row = 1
    Do While Not rs.EOF
        With ClaimsInputV2.ClaimsListBox
            .AddItem
            .List(row, 0) = rs.Fields("BILL_DIV_ID").Value 'Division-
            .List(row, 1) = rs.Fields("OFR_NBR").Value 'offer-
            .List(row, 2) = rs.Fields("AMOUNT").Value    'amount-
            .List(row, 3) = rs.Fields("VDC_DATE").Value 'vdc date
            .List(row, 4) = rs.Fields("VDC_NUMBER").Value   'vdc number
            .List(row, 5) = rs.Fields("BILLER_ERROR").Value 'biller error-
            .List(row, 6) = rs.Fields("TRIBBLE").Value 'tribble-
            .List(row, 7) = rs.Fields("PERIOD").Value
            .List(row, 8) = rs.Fields("YEAR").Value 'year
            .List(row, 9) = rs.Fields("BILL_DIV_ID").Value & "--" & rs.Fields("VEND_ACCT_NBR").Value & "--" & rs.Fields("LOG").Value & "--" & rs.Fields("VEND_BILL_NM").Value & _
            "--" & rs.Fields("ALLOW_TYPE").Value & "--" & rs.Fields("PAY_TYPE").Value & "--" & rs.Fields("PERFORM_1").Value & "--" & rs.Fields("PERFORM_2").Value & "--" & _
            "NULL" & "--" & rs.Fields("SF_NBR").Value & "--" & rs.Fields("CCOA_ACCT_NBR").Value & "--" & "NULL" & "--" & rs.Fields("SF_NBR2").Value & "--" & rs.Fields("CCOA_ACCT_NBR2").Value & _
            "--" & "NULL" & "--" & rs.Fields("VDC_NUMBER").Value & "--" & rs.Fields("VDC_DATE").Value & "--" & rs.Fields("CURR_DATE").Value & "--" & "NULL"
        row = row + 1
        rs.MoveNext
        End With
    Loop
    
'Error Handling
    Exit Sub
catch:
    Console.error err.Description, "Claims.FillCABSData"
    DisplayErrorMessage "Claims.FillCABSData" & vbCr & vbCr & err.Description
    
End Sub
'Created 10/26/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose: This function is used to call the row of data from DB2 and popualate
'the info along with the data from the form into the .txt file to be inputted into SF
'References:DB2.RunQuery,ReadTextFile,checkIfNull,ClaimsInput (Forms),clearExtraSpaces
Public Function MakeClaimQuery_V2() As String
On Error GoTo catch
    Dim SQL As String
    Dim DMLogic As String
    Dim invoice As String
    Dim TribbleID As String
    
    'DEBIT MEMO LOGIC
    'CABS LOGIC
    'OFFER NUMBER
    If InStr(ClaimsInputV2.OfferNumTextBox.Value, "D") > 0 Then
        DMLogic = "0"
        invoice = ClaimsInputV2.OfferNumTextBox.Value
    ElseIf Replace(ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 4), " ", "") <> "" Then
        DMLogic = ClaimsInputV2.OfferNumTextBox.Value
        invoice = ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 4)
    Else
        DMLogic = ClaimsInputV2.OfferNumTextBox.Value
        invoice = "NULL"
    End If
    
    'Get the TribbleID '01/07/2022 - might be activated again in the future
    'If ClaimsInputV2.TribbleComboBox = "True" Then
    '    TribbleID = InputBox("Please Enter the Tribble ID for this Claims", "Tribble ID Inquiry")
    'Else
    TribbleID = "00000"
    'End If
    
    'VENDOR_NAME
    Dim DeleteApost As String 'replace apostrophes
    DeleteApost = GetDB2SelectedClaim(VENDOR_NAME)
    DeleteApost = Replace(ClearExtraSpaces(GetDB2SelectedClaim(VENDOR_NAME)), "'", "")
    
    SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Input_V3.txt")
    
    SQL = ClaimsSQLStringReplace_V2(SQL, GetLastClaimID + 1, GetDB2SelectedClaim(DIV), checkIfNull(GetDB2SelectedClaim(Vendor)), GetDB2SelectedClaim(Log), DMLogic, DeleteApost, _
    ClaimsInputV2.AmountTextBox.Value, ClaimsInputV2.ReasonComboBox.Value, ClaimsInputV2.BillerErrorComboBox.Value, ClaimsInputV2.TribbleComboBox.Value, _
    GetDB2SelectedClaim(ALLOW_TYPE), GetDB2SelectedClaim(PAY_TYPE), GetDB2SelectedClaim(PERFORM_1), GetDB2SelectedClaim(PERFORM_2), checkIfNull(GetDB2SelectedClaim(FAC_NBR)), _
    checkIfNull(GetDB2SelectedClaim(SF_NBR)), checkIfNull(GetDB2SelectedClaim(CCOA_ACCT_NBR)), checkIfNull(GetDB2SelectedClaim(FAC_NBR2)), checkIfNull(GetDB2SelectedClaim(SF_NBR2)), checkIfNull(GetDB2SelectedClaim(CCOA_ACCT_NBR3)), _
    GetDB2SelectedClaim(MERCHANDISER), invoice, ClaimsInputV2.PeriodComboBox.Value, ClaimsInputV2.YearComboBox.Value, GetDB2SelectedClaim(VDC_DTE), GetDB2SelectedClaim(CURRENT_DATE), UCase(ClaimsInputV2.UsernameComboBox.Value), "", TribbleID)
    'Debug.Print SQL
    MakeClaimQuery_V2 = SQL
    
    'Error Handling
    Exit Function
catch:
    Console.error err.Description, "Claims.MakeClaimQuery_V2"
    DisplayErrorMessage "Claims.MakeClaimQuery_V2" & vbCr & vbCr & err.Description
    
End Function

'Created 10/26/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose: This is a fun popup box that inspects he tolerances of the amount and type of claim submitted
'If the tolerance is met then surprise random picture is generated
Public Sub ClaimsTolerancesPopUpPicture_V2(ByRef AllowType As String, ByRef Data As Scripting.Dictionary)
    'Fun Pop-up For Amount greater than tolerance
    If (AllowType = "C" And (CDec(ClaimsInputV2.AmountTextBox.Value) >= CDec(Data("ToleranceC")))) _
        Or (AllowType = "A" And (CDec(ClaimsInputV2.AmountTextBox.Value) >= CDec(Data("ToleranceA")))) _
        Or (AllowType = "S" And (CDec(ClaimsInputV2.AmountTextBox.Value) >= CDec(Data("ToleranceS")))) _
        Or (AllowType = "T" And (CDec(ClaimsInputV2.AmountTextBox.Value) >= CDec(Data("ToleranceT")))) _
        Or (AllowType = "SEF" And (CDec(ClaimsInputV2.AmountTextBox.Value) >= CDec(Data("ToleranceSEF")))) _
        Or (AllowType = "FCF" And (CDec(ClaimsInputV2.AmountTextBox.Value) >= CDec(Data("ToleranceFCF")))) _
        Or (AllowType = "VCF" And (CDec(ClaimsInputV2.AmountTextBox.Value) >= CDec(Data("ToleranceVCF")))) _
        Or (AllowType = "FTF" And (CDec(ClaimsInputV2.AmountTextBox.Value) >= CDec(Data("ToleranceFTF")))) Then
            SurpriseFormRandomPicture
    End If
End Sub
Public Function ValidateClaimsFieldsNotEmpty_V2() As Boolean
    With ClaimsInputV2
        If .OfferNumTextBox.Value = vbNullString Or _
        .AmountTextBox.Value = vbNullString Or _
        .ReasonComboBox.Value = vbNullString Or _
        .BillerErrorComboBox.Value = vbNullString Or _
        .TribbleComboBox.Value = vbNullString Then
            ValidateClaimsFieldsNotEmpty_V2 = False
            MsgBox "One of the TextBoxes is Empty"
        Else
            ValidateClaimsFieldsNotEmpty_V2 = True
        End If
    End With
End Function

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 04/02/2021
'Purpose:This is a temporary function that will be used to handle the CABS invoice until a query can be designed to integrate into the process.
'All of the usual fields are pulled and inputed into snowflake, however all of the accounting units will be pulled in later with a query
'This process is accessed by selecting Manual Input at the top of the tool.
Public Sub MaunalInputFunction_V2()
On Error GoTo catch
    Dim DIV As String
    Dim SQL As String
    Dim timestamp As String
    Dim datestamp As String
    
    timestamp = " " & CStr(Hour(Now)) & ":" & CStr(Minute(Now)) & ":" & CStr(Second(Now)) 'hr:mm:ss"
    datestamp = CStr(YEAR(Now)) & "-" & CStr(Month(Now)) & "-" & CStr(Day(Now))
    
    'Input for Division Function
    DIV = InputBox("DIVISION: ", "Manual Input")
    
    'Build SQL query
    SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Input_V3.txt")
    SQL = ClaimsSQLStringReplace_V2(SQL, GetLastClaimID + 1, DIV, 0, 0, _
    ClaimsInputV2.OfferNumTextBox.Value, "NULL", ClaimsInputV2.AmountTextBox.Value, ClaimsInputV2.ReasonComboBox.Value, ClaimsInputV2.BillerErrorComboBox.Value, _
    ClaimsInputV2.TribbleComboBox.Value, "NULL", "NULL", "NULL", "NULL", "NULL", "NULL", "NULL", "NULL", "NULL", "NULL", "NULL", 0, ClaimsInputV2.PeriodComboBox.Value, _
    ClaimsInputV2.YearComboBox.Value, datestamp, datestamp, ClaimsInputV2.UsernameComboBox.Value, "")
             
    InsertDataIntoSnowflake (SQL) 'SQL Query
    
    MsgBox "The Claim below has been manually submitted: " & vbCr & _
    ClaimsInputV2.OfferNumTextBox.Value & vbTab & "Amount: $" & ClaimsInputV2.AmountTextBox.Value
    
    'check if archiving is needed
    'CheckIfArchiveIsReady_V2
    
    ClaimsInputV2.ClaimsListBox.Clear

    Exit Sub
catch:
    Dim ErrorMessage As String
    ErrorMessage = err.Description

    Application.ScreenUpdating = True
    Console.error ErrorMessage, "Claims.CABS_TEMP_FUNCTION"
    
    DisplayErrorMessage ErrorMessage
    
End Sub

'Created 10/28/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose:The purpose of this function is to see if a claim exists. If a claim does exist then
'True is returned Else False.
'References:makeClaimSearchQuery
Public Function CheckIfClaimExists_V2(ByVal DIV As String, ByVal OfferNumber As String, ByVal Amount As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim msgString As String
    Dim answer As String
    Dim SQL As String
    
    If InStr(UCase(OfferNumber), "D") > 0 Then
        SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Delete_Search_DM_SF.txt")
        SQL = ClaimsSQLStringReplace_V2(SQL, "", DIV, , , , , Amount, "", "", "", "", "", "", "", "", "", "", "", "", "", "", OfferNumber, , , , , , "")
    Else
        SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Delete_Search_SF.txt")
        SQL = ClaimsSQLStringReplace_V2(SQL, "", DIV, , , OfferNumber, , Amount, , , , , , , , , , , , , , , , , , , , , "")
    End If
    
    Set rs = QuerySnowFlake(SQL)
    
    If IsRecordsetEmpty(rs) Then
        CheckIfClaimExists_V2 = False
        Exit Function
    End If
    
    msgString = "This amount has already been claimed for this division, please verify that this is not a duplicate of the below claim before proceeding. " & vbCr
    
    Dim DataString As String
    Do While Not rs.EOF
            msgString = msgString + "Amount: $" & rs.Fields("AMOUNT").Value & vbTab & _
            "User: " & rs.Fields("CLAIMED").Value & vbTab & _
            "Date: " & rs.Fields("DATE_CREATED").Value & vbCr & _
            "ID#: " & rs.Fields("CLAIM_ID").Value & vbCr
            
        rs.MoveNext
    Loop
    
    msgString = msgString + "Press [OK] to Submit and [Cancel] to Exit"
    answer = MsgBox(msgString, vbQuestion + vbOKCancel, "User Repsonse")
    
    If answer = vbOK Then
        CheckIfClaimExists_V2 = False
    Else
        CheckIfClaimExists_V2 = True
    End If
    
End Function

'Created 10/28/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose:This function updates the claims that are pulled from Snowflake.
'Only the options referenced below can be changed currently
'References:InsertDataIntoSnowflake
Public Sub UpdateClaimsInSF_V2()
    Dim SQL As String
    Dim ClaimType As String
    Dim err As Integer
    Dim OfferNumber As String
    
    'Logic for DEBIT MEMO
    If InStr(ClaimsInput.OfferNumTextBox.Value, "D") > 0 Then
        SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Update_SF_DebitMemo.txt")
        ClaimType = "Debit Memo: "
    Else
        SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Update_SF.txt")
        ClaimType = "Offer Number: "
    End If

    SQL = ClaimsSQLStringReplace_V2(SQL, GetClaimIDFromSelection, ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 0), , , ClaimsInputV2.OfferNumTextBox.Value, , _
    ClaimsInputV2.AmountTextBox.Value, ClaimsInputV2.ReasonComboBox.Value, ClaimsInputV2.BillerErrorComboBox.Value, _
    ClaimsInputV2.TribbleComboBox.Value, , , , , , , , , , , , , ClaimsInputV2.PeriodComboBox.Value, ClaimsInputV2.YearComboBox.Value, _
    , , UCase(CStr(ClaimsInputV2.UsernameComboBox.Value)), ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 2), "00000")

    err = InsertDataIntoSnowflake(SQL)
    If err = 0 Then
        MsgBox "The Claim below has been updated:" & vbCr & ClaimType & _
        ClaimsInputV2.OfferNumTextBox.Value & vbTab & "Amount: $" & ClaimsInputV2.AmountTextBox.Value
    Else
        MsgBox "An Error Occured Updating The Claim. Please contact BA"
    End If
    
    'Reset the Form
    ClearClaimFormBoxes_V2
    
End Sub
Public Sub DeleteClaimFromSF_V2()
    Dim SQL As String

    SQL = "DELETE FROM DW_PRD.TEMP_TABLES.PA_CLAIMS_CLONE WHERE CLAIM_ID = " & CStr(GetClaimIDFromSelection)
    err = InsertDataIntoSnowflake(SQL)
    If err = 0 Then
        MsgBox "The Claim below has been Deleted"
    Else
        MsgBox "An Error Occured While Deleting The Claim. Please contact BA"
    End If
    
End Sub
Public Function GetClaimIDFromSelection() As Long
    Dim SplitString() As String
    SplitString = split(CStr(ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 9)), "---")
    GetClaimIDFromSelection = SplitString(1)
End Function
Public Function GetDB2SelectedClaim(SelectedValue As claimsDB2Information) As Variant
    Dim Data() As String
    Data = split(CStr(ClaimsInputV2.ClaimsListBox.List(ClaimsSelectionIndex_V2, 9)), "--")
    'Debug.Print SelectedValue
    'GetDB2SelectedClaim = SelectedValue
    GetDB2SelectedClaim = Data(SelectedValue)
End Function
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 12/02/2021
'Purpose: To retrieve the last TribbleID that exists in the DB
'References:ReadTextFile,IsRecordsetEmpty QuerySnowFlake
Public Function GetLastClaimID() As Long
    Dim SQL As String
    Dim rs As ADODB.Recordset
    SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Get_Last_Claims.txt") 'find the last tribbler

    Set rs = QuerySnowFlake(SQL)
    
    If IsRecordsetEmpty(rs) Then
        Exit Function
    End If
    
    Do While Not rs.EOF
        GetLastClaimID = rs.Fields("CLAIM_ID").Value
    rs.MoveNext
    Loop
    
End Function
Public Function ValidateClaimsLookUpField_V2()
        'convert lowercase d -> D
    If InStr(ClaimsInputV2.OfferNumTextBox.Value, "d") > 0 Then
        ClaimsInputV2.OfferNumTextBox.Value = Replace(ClaimsInputV2.OfferNumTextBox.Value, "d", "D")
    End If
End Function
'TODO ADD MORE REGUALR EXPRESSIONS
Public Function ValidateClaimsFieldsFormatting_V2() As Boolean
    
    'Regular Expressions to check the values in the form fields
    If regularExpressionExists(AmountTextBox.Value, "^\$?[0-9]+\.?[0-9]*$") = False Then
        MsgBox "ONLY DIGITS & TWO FLOATS FOR AMOUNT EX. 1234.00"
        ValidateClaimsFieldsFormatting = False
    Else
        ValidateClaimsFieldsFormatting = True
    End If
    
End Function
'Created 06/17/2021
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose:
'References:
Public Sub ClearClaimFormBoxes_V2()
    'Clear the Claim Form
    ClaimsInputV2.AmountTextBox.Value = ""
    ClaimsInputV2.ReasonComboBox.Value = ""
    ClaimsInputV2.BillerErrorComboBox.Value = ""
    ClaimsInputV2.TribbleComboBox.Value = ""
    ClaimsInputV2.PeriodComboBox.Value = ""
    ClaimsInputV2.YearComboBox.Value = ""
    ClaimsInputV2.UsernameComboBox.Value = Environ("Username")
    'ClaimsInputV2.TribbleTextBox.value = ""
    ClaimsInputV2.IDTextBox.Value = ""
    ClaimsInputV2.ClaimsListBox.Clear
End Sub
Public Sub InitalizeReasons_V2()
On Error GoTo catch
    Dim Data As Scripting.Dictionary
    Dim item As Variant
    Set Data = JsonParse(ReadTextFile(ENV.use("CONFIG_CLAIMS")))
    
    ClaimsInputV2.ReasonComboBox.Clear
    
    With ClaimsInputV2.ReasonComboBox
        If ClaimsInputV2.ClaimType.Value = "Debit Memo" Then
            For Each item In Data("reasons_wims")
                item = Replace(item, "Â", "")
                .AddItem item
            Next item
        Else
            For Each item In Data("reasons_pacs")
                item = Replace(item, "Â", "")
                .AddItem item
            Next item
        End If
    End With
    
    'Error Handling
    Exit Function
catch:
    Console.error err.Description, "Claims.InitalizeReasons_V2"
    DisplayErrorMessage "Claims.InitalizeReasons_V2" & vbCr & vbCr & err.Description
End Sub
'Created 10/26/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose:To check to see if a InputString is null. If it is Null
'First it converts all white space values to 0 and then check to see
'if the datavalue is missing. If it is then the NULL is attached
'References:
Private Function checkIfNull(InputString As String) As String
    On Error GoTo catch

    InputString = Replace(InputString, " ", "")
    If InputString = "" Then
        checkIfNull = "NULL"
    Else
        checkIfNull = InputString
    End If
    
    Exit Function
catch:
    Dim ErrorMessage As String
    ErrorMessage = err.Description

    Application.ScreenUpdating = True
    Console.error ErrorMessage, "Claims.checkIfNull"
    
    DisplayErrorMessage ErrorMessage
End Function
'Created: 06/16/2021
'Purpose:
Public Function IdentifyClaimsType_V2() As String
On Error GoTo catch
    Dim REGEXCabsInvoice As String
    REGEXCabsInvoice = "\b[A-Z]{3}\d{9,10}-\d{2}P{0,1}\b"
    
    If InStr(ClaimsInputV2.OfferNumTextBox.Value, "D") > 0 Then
        IdentifyClaimsType_V2 = "DM"
    ElseIf regularExpressionExists(CStr(ClaimsInputV2.OfferNumTextBox.Value), REGEXCabsInvoice) = True Then
        IdentifyClaimsType_V2 = "CABS"
    Else
        IdentifyClaimsType_V2 = "OFFER"
    End If
    
    'Error Handling
    Exit Function
catch:
    Console.error err.Description, "Claims.IdentifyClaimsType_V2"
    DisplayErrorMessage "Claims.IdentifyClaimsType_V2" & vbCr & vbCr & err.Description
End Function
'Created 10/26/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose:
Public Function ClaimsSelectionIndex_V2() As Integer
    Dim index As Integer
    For index = 0 To ClaimsInputV2.ClaimsListBox.ListCount - 1
        If ClaimsInputV2.ClaimsListBox.Selected(index) = True And index <> 0 Then
            ClaimsSelectionIndex_V2 = index
            Exit Function
        End If
    Next index
End Function

'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 03/24/21
'Purpose: Possible reduction in code for populating a text file string with values
'Current issue is the same of the sql files do not contain the same replace variables
Public Function ClaimsSQLStringReplace_V2(SQL As String, CLAIM_ID As String, DIV As String, Optional VENDOR_NUMBER As String, Optional LOG_NUMBER As String, _
Optional OFFER_NUMBER As String, Optional VENDOR_NAME As String, Optional Amount As String, _
Optional REASON As String, Optional BILLER_ERROR As String, Optional TRIBBLE As String, Optional ALLOW_TYPE As String, Optional PAY_TYPE As String, Optional PERFORM_1 As String, _
Optional PERFORM_2 As String, Optional FAC_NBR As String, Optional SF_NBR As String, Optional CCOA_ACCT_NBR As String, Optional FAC_NBR2 As String, Optional SF_NBR2 As String, _
Optional CCOA_ACCT_NBR3 As String, Optional MERCHANDISER As String, Optional VDC_NBR As String, Optional period As String, Optional YEAR As String, Optional VDC_DTE As String, _
Optional DATE_CREATED As String, Optional CLAIMED As String, Optional AMOUNTold As String, Optional TribbleID As String) As String
    On Error GoTo catch
    SQL = Replace(SQL, "(&CLAIM_ID)", CLAIM_ID)
    SQL = Replace(SQL, "(&DIV)", DIV)
    SQL = Replace(SQL, "(&VENDOR_NUMBER)", VENDOR_NUMBER)
    SQL = Replace(SQL, "(&LOG_NUMBER)", LOG_NUMBER)
    SQL = Replace(SQL, "(&OFFER_NUMBER)", OFFER_NUMBER)
    SQL = Replace(SQL, "(&VENDOR_NAME)", VENDOR_NAME)
    SQL = Replace(SQL, "(&AMOUNT)", Amount)
    SQL = Replace(SQL, "(&REASON)", REASON)
    SQL = Replace(SQL, "(&BILLER_ERROR)", BILLER_ERROR)
    SQL = Replace(SQL, "(&TRIBBLE)", TRIBBLE)
    SQL = Replace(SQL, "(&ALLOW_TYPE)", ALLOW_TYPE)
    SQL = Replace(SQL, "(&PAY_TYPE)", PAY_TYPE)
    
    'check if NULL
    SQL = Replace(SQL, "(&PERFORM_1)", PERFORM_1)
    SQL = Replace(SQL, "(&PERFORM_2)", PERFORM_2)
    SQL = Replace(SQL, "(&FAC_NBR)", FAC_NBR)
    SQL = Replace(SQL, "(&SF_NBR)", SF_NBR)
    SQL = Replace(SQL, "(&CCOA_ACCT_NBR)", CCOA_ACCT_NBR)
    SQL = Replace(SQL, "(&FAC_NBR2)", FAC_NBR2)
    SQL = Replace(SQL, "(&SF_NBR2)", SF_NBR2)
    SQL = Replace(SQL, "(&CCOA_ACCT_NBR3)", CCOA_ACCT_NBR3)
    SQL = Replace(SQL, "(&MERCHANDISER)", MERCHANDISER)
    SQL = Replace(SQL, "(&VDC_NBR)", VDC_NBR)
    'End of Check with Null
    
    SQL = Replace(SQL, "(&PERIOD)", period)
    SQL = Replace(SQL, "(&YEAR)", YEAR)
    SQL = Replace(SQL, "(&VDC_DTE)", VDC_DTE)
    SQL = Replace(SQL, "(&DATE_CREATED)", DATE_CREATED)
    SQL = Replace(SQL, "(&CLAIMED)", CLAIMED)
    
    SQL = Replace(SQL, "(&AMOUNTold)", AMOUNTold)
    SQL = Replace(SQL, "(&TRIBBLEID)", TribbleID)

    ClaimsSQLStringReplace_V2 = SQL
    
'Error Handling
    Exit Function
catch:
    Console.error err.Description, "Claims.ClaimsSQLStringReplace_V2"
    DisplayErrorMessage "Claims.ClaimsSQLStringReplace_V2" & vbCr & vbCr & err.Description
End Function

'
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////ANALYTICS FOR CLAIMS ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'
'Created: 12/17/2021
'Author: Nicholas Ackerman
'Purpose:
Public Sub GetClaimsAnalytics()
On Error GoTo catch
    Dim SQL As String
    Dim rs As ADODB.Recordset
    
    'Run the Query Process
    SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Analytics.txt")
    SQL = Replace(SQL, "(&USER)", ClaimsAnalytics.UsernameComboBox.Value)
    SQL = Replace(SQL, "(&PERIOD)", ClaimsAnalytics.PeriodComboBox.Value)
    SQL = Replace(SQL, "(&YEAR)", ClaimsAnalytics.YearComboBox.Value)

    Set rs = QuerySnowFlake(SQL)
    
    If IsRecordsetEmpty(rs) Then
        Exit Sub
    End If
    
    Dim Count As Integer
    Count = 1
    Do While Not rs.EOF
        If Count = 1 Then
            ClaimsAnalytics.TotalInfoBoxPeriod = rs.Fields("DATA").Value
        ElseIf Count = 2 Then
            ClaimsAnalytics.TotalInfoBoxYear = rs.Fields("DATA").Value
        ElseIf Count = 3 Then
            ClaimsAnalytics.TotalInfoBox = rs.Fields("DATA").Value
        ElseIf Count = 4 Then
            ClaimsAnalytics.UserInfoBoxPeriod = rs.Fields("DATA").Value
        ElseIf Count = 5 Then
            ClaimsAnalytics.UserInfoBoxYear = rs.Fields("DATA").Value
        ElseIf Count = 6 Then
            ClaimsAnalytics.UserInfoBox = rs.Fields("DATA").Value
        End If
        Count = Count + 1
    rs.MoveNext
    Loop

'Error Handling
    Exit Sub
catch:
    Console.error err.Description, "Claims.GetClaimsAnalytics"
    DisplayErrorMessage "Claims.GetClaimsAnalytics" & vbCr & vbCr & err.Description

End Sub
'Created: 12/17/2021
'Author: Nicholas Ackerman
'Purpose:
Public Sub GetClaimsAnalyticsResults()
On Error GoTo catch
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Dim Count As Integer
    
    'Run the Query Process
    SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Analytics_Results.txt")
    SQL = Replace(SQL, "(&USER)", ClaimsAnalytics.UsernameComboBox.Value)
    SQL = Replace(SQL, "(&PERIOD)", ClaimsAnalytics.PeriodComboBox.Value)
    SQL = Replace(SQL, "(&YEAR)", ClaimsAnalytics.YearComboBox.Value)

    Set rs = QuerySnowFlake(SQL)
    
    If IsRecordsetEmpty(rs) Then
        Exit Sub
    End If
    
        With ClaimsAnalytics.ListBox1
        .AddItem
        .List(0, 0) = "CLAIM_ID"
        .List(0, 1) = "DIV"
        .List(0, 2) = "OFFER_NUMBER"
        .List(0, 3) = "AMOUNT"
        .List(0, 4) = "VDC_NBR"
    End With
    
    Count = 1
    Do While Not rs.EOF
        With ClaimsAnalytics.ListBox1
            .AddItem
            .List(Count, 0) = rs.Fields("CLAIM_ID").Value
            .List(Count, 1) = rs.Fields("DIV").Value
            .List(Count, 2) = rs.Fields("OFFER_NUMBER").Value
            .List(Count, 3) = rs.Fields("AMOUNT").Value
            .List(Count, 4) = rs.Fields("VDC_NBR").Value
        End With
        Count = Count + 1
        rs.MoveNext
    Loop
    
'Error Handling
    Exit Sub
catch:
    Console.error err.Description, "Claims.GetClaimsAnalyticsResults"
    DisplayErrorMessage "Claims.GetClaimsAnalyticsResults" & vbCr & vbCr & err.Description
End Sub
'Created: 12/17/2021
'Author: Nicholas Ackerman
'Purpose: Ensure that all three input fields are entered
Public Function ClaimsAnalyticsValidation() As Boolean
    If ClaimsAnalytics.UsernameComboBox.Value <> "" And ClaimsAnalytics.YearComboBox.Value <> "" And ClaimsAnalytics.PeriodComboBox.Value <> "" Then
        ClaimsAnalyticsValidation = True
    Else
        ClaimsAnalyticsValidation = False
    End If
End Function

'Created: 12/17/2021
'Author: Nicholas Ackerman
'Purpose: Open the selected claim from the analytics window
Public Sub OpenClaimsIDFromAnalytics()
On Error GoTo catch
    Dim ClaimsID As String
    Dim SQL As String
    Dim index As Integer
    
    For index = 0 To ClaimsAnalytics.ListBox1.ListCount - 1
        If ClaimsAnalytics.ListBox1.Selected(index) = True And index <> 0 Then
             ClaimsID = ClaimsAnalytics.ListBox1.List(index, 0)
            Exit For
        End If
    Next index
    
    ClaimsAnalytics.Hide
    ClearClaimFormBoxes_V2

    SQL = ReadTextFile(ENV.use("SQLFOLDERPATH") & "\Claims\Claims_V2\Claims_Search_ClaimsID.txt")
    SQL = Replace(SQL, "(&CLAIMS_ID)", ClaimsID)
    
    LocateClaimsInSF_V2 SQL

'Error Handling
    Exit Sub
catch:
    Console.error err.Description, "Claims.OpenClaimsIDFromAnalytics"
    DisplayErrorMessage "Claims.OpenClaimsIDFromAnalytics" & vbCr & vbCr & err.Description
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 01/03/2022
'Purpose:The purpose of this function is archive the claims into a .csv file automatically.
'This will allow for us to have a backup incase the database ever goes down. Currently this
'Function is set to archive every 50 claims.
'TODO CHANGE THE ARCHIVE FUNCTION TO PULL IN THE DATA FROM THE NEW DATABASE
Public Sub CheckIfArchiveIsReady_V2()
    Dim rs As ADODB.Recordset
    Dim ROW_NUM As Integer
    
    Set rs = QuerySnowFlake("SELECT COUNT(*) AS ROW_NUM FROM DW_PRD.TEMP_TABLES.PA_CLAIMS_CLONE")
    Do While Not rs.EOF
        ROW_NUM = rs.Fields("ROW_NUM").Value
        rs.MoveNext
    Loop
    
    If ROW_NUM Mod 50 = 0 Then
        Dialog.show "Archiving Process Started", "DO NOT CLOSE OUT OF WINDOWS"
        ArchiveSFDataBases "Claims"
        Dialog.Hide
    End If
    
End Sub
'Author: Nicholas Ackerman @<nicholas.ackerman@albertsons.com>
'Created: 09/17/2020
'Purpose:This is called after pressing the button ton the main screen of RT amcros.
'References:GetRawTribbleDataReport
Public Sub ClaimsReporting()
    On Error GoTo catch
    'On Error Resume Next

    Dim wb As Workbook
    Set wb = Workbooks.Open("K:\AA\SHARE\AuditTools\rtmacros\data\Reports\Claims_Report.xlsm")

'Error Handling
    Exit Sub
catch:
    Console.error err.Description, "Claims.ClaimsReporting"
    DisplayErrorMessage "Claims.ClaimsReporting" & vbCr & vbCr & err.Description
End Sub
'Created 06/17/2021
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose:
'References:
Public Sub ClearClaimFormBoxesV2()
    'Clear the Claim Form
    ClaimsInputV2.AmountTextBox.Value = ""
    ClaimsInputV2.ReasonComboBox.Value = ""
    ClaimsInputV2.BillerErrorComboBox.Value = ""
    ClaimsInputV2.TribbleComboBox.Value = ""
    ClaimsInputV2.PeriodComboBox.Value = ""
    ClaimsInputV2.IDTextBox.Value = ""
    ClaimsInputV2.YearComboBox.Value = ""
    ClaimsInputV2.UsernameComboBox.Value = Environ("Username")
    ClaimsInputV2.ClaimsListBox.Clear
End Sub


