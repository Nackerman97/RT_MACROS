Attribute VB_Name = "APIFunctions"
Option Explicit
Public Enum CABS_INTERFACE
    UAT = 1
    PROD = 2
End Enum
Public Function CABSENV(ByRef Ind As Integer)
    CABSENV = Choose(Ind, "LINK NEEDED","LINK NEEDED")
End Function
'Created 08/16/2021 BY Nicholas Ackerman
'Purpose: The purpose of this function is include a list of SUBS that will pull necessary information for functions across RT_MACROS.
'This particular sub is used to pull the BILLING RECORD information from CABS
'REFERENCES: printJSAONDATA
'CABSGetBRNopaInfo "nacke08", "10731144"
Public Sub CABSGetBRNopaInfo(ByVal Username As String, ByVal BillingRecord As String, conn As CABS_INTERFACE)
    Dim Req As New MSXML2.XMLHTTP60
    Dim returnData, URL As String
    
    URL = CABSENV(CInt(conn)) & "/pabsc/billingrecord/nopaInfo/" & BillingRecord
    
    With Req
        .Open "GET", URL, False  'GET BILLING RECORD
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "x-auth-user", UCase(Username)
        .Send
        returnData = .responseText
    End With
    
    printJSONDATA (returnData)
    
End Sub
'Created 08/16/2021 BY Nicholas Ackerman
'Purpose: The purpose of this function is include a list of SUBS that will pull necessary information for functions across RT_MACROS.
'This particular sub is used to pull the OFFEER_NUMBER and DIV information from CABS. This will pull in the
'CABSGetOfferInfo "NACKE08", "33", "4854594"
'REFERENCES: printJSAONDATA
Public Sub CABSGetOfferInfo(ByVal Username As String, ByVal DIV As String, ByVal OfferNumber As String, conn As CABS_INTERFACE)
    Dim Req As New MSXML2.XMLHTTP60
    Dim returnData, URL As String
    
    URL = CABSENV(CInt(conn)) & "/pabsc/autoBiller/offerData/PRE/" & Username & "/" & DIV & "/" & OfferNumber
    
    With Req
        .Open "GET", URL, False
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "x-auth-user", UCase(Username)
        .Send
        returnData = .responseText
    End With
    
    'printJSONDATA (returnData)
    printJSONDATAtoActiveSheet (returnData)
    
End Sub
'Created 08/16/2021 BY Nicholas Ackerman
'Purpose:
Public Sub printJSONDATA(ByVal returnData As String)
    Dim Data As New Collection
    Dim item As Scripting.Dictionary
    Set Data = ParseJson(returnData)
    
    On Error GoTo catch

    Dim key As Variant
    For Each item In Data
        For Each key In item.Keys
            If key <> "indicator" Then
                Debug.Print key, item(key)
            End If
        Next key
        Debug.Print "-------- NEW RECORD"
    Next item

catch:
    
End Sub
'Example to show how the to create a string of JSON data
Public Sub printJSONDATATESTING()
    
    Dim P1 As New Scripting.Dictionary
    P1("name") = "Nick"
    P1("age") = 24
    
    Dim P2 As New Scripting.Dictionary
    P2("name") = "Robert"
    P2("age") = 33
    
    Dim people As New Collection
    people.Add P1
    people.Add P2
    
    'Place this into a collection
    Debug.Print JsonStringify(people)

End Sub
Public Sub printJSONDATAtoActiveSheet(ByVal returnData As String)
    Dim Data As New Collection
    Dim item As Scripting.Dictionary
    Dim row, col As Long
    Set Data = ParseJson(returnData)
    
    row = 2
    col = 1

    Dim key As Variant
    For Each item In Data
        col = 1
        For Each key In item.Keys
            Cells(1, col).Value = key
            Cells(row, col).Value = item(key)
            col = col + 1
        Next key
        row = row + 1
    Next item
End Sub
'//////////////////////////////////////////////////////////////////
'CLAIMS TOOL DESIGN
'The purpose of this tool will be to pull claims based off of a user input of an offer number
'Will need to first determine if the offer number resides in CABS or in DB2
'The offer in CABS will need to move throught the API to pull in information that will lead to the billing record.
'The billing record will contain the claim.
'
Public Sub ClaimsCABSBillingRecord(ByRef DIV As String, ByRef OFFER_NUMBER As String, ByRef BR As String)
'FIELDS NEEDED DIV, OFFER NUMBER, AMOUNT, INVOICE DATE, INVOICE NUMBER, OFFER NUMBER, PERIOD, YEAR,
    Dim INVOICE_NBR, Vendor, Log, AllowType, SUM_AMOUNT, PerfCode1, PerfCode2, VendorName  As String
    'BR = "10890326"   'HDR = 4740698
    'OFFER_NUMBER = 4912061
    'DIV = 20
    
    Dim Req As New MSXML2.XMLHTTP60
    Dim returnData, URL As String
    
    URL = CABSENV(2) & "/pabsc/autoBiller/offerData/PRE/" & "NACKE08" & "/" & DIV & "/" & OFFER_NUMBER
    With Req
        .Open "GET", URL, False
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "x-auth-user", UCase("NACKE08")
        .Send
        returnData = .responseText
    End With
    
    Dim Data As New Collection
    Dim item As Scripting.Dictionary
    Dim DataString() As String

    Set Data = ParseJson(returnData)
    
    'pull the offer number and the vendor number
    Dim key As Variant
    For Each item In Data
        For Each key In item.Keys
            If key = "vendor" Then
                Vendor = CStr(item(key))
            ElseIf key = "log" Then
                Log = CStr(item(key))
            ElseIf key = "type" Then
                AllowType = CStr(item(key))
            End If
        Next key
    Next item
    
    

'/////////////////////////////////////STEP 2////////////////////////////////// billing record header

    URL = CABSENV(2) & "/pabsc/billingrecord/calculateamnt/" & BR
    With Req
        .Open "GET", URL, False  'GET BILLING RECORD
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "x-auth-user", UCase("NACKE08")
        .Send
        returnData = .responseText
    End With
    
    Set Data = ParseJson("[" & returnData & "]")

    For Each item In Data
        For Each key In item.Keys
            If key = "totalSumIncomeAmount" Then
                SUM_AMOUNT = CStr(item(key))
            End If
        Next key
    Next item


'////////////////////////////////////STEP 3 /////////////////////////////

    URL = CABSENV(2) & "/pabsc/billingrecord/nopaInfo/" & BR
    With Req
        .Open "GET", URL, False  'GET BILLING RECORD
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "x-auth-user", UCase("NACKE08")
        .Send
        returnData = .responseText
    End With

    Set Data = ParseJson(returnData)

    'pull the offer number and the vendor number
    For Each item In Data
        For Each key In item.Keys
            If key = "perfCdOne" Then
                PerfCode1 = CStr(item(key))
            ElseIf key = "perfCdTwo" Then
                PerfCode2 = CStr(item(key))
            ElseIf key = "manfName" Then
                VendorName = CStr(item(key))
            End If
        Next key
    Next item
    
    
    'Use a query to pull in the rest of the infomration. Use the CABS API to pull in the billed amount information
    
    'DIV, VENDOR, LOG, offer, vendor name, amount, reason, biller error, tribble, allow tyoe,
    Debug.Print DIV & "-" & Vendor & "-" & Log & "-" & OFFER_NUMBER & "-" & VendorName & "-" & SUM_AMOUNT & "-" & AllowType & "-" & PerfCode1 & "-" & PerfCode2
    
End Sub


