Attribute VB_Name = "CABS"
Option Explicit


'!!!!!!!!!!!!!!!!!'TRUE IS CABS and FALSE IS PACS
'Purpose: The purpose of this function is to take in two inputs of an offernumber and a
'divison. A query is then run and the result is either CABS or PACS for billing
'location. If True then in CABS else false is in PACS
'References:ReadTextFile, replace, QuerySnowFlake
'Author: Nicholas Ackerman
'Created: 01-19-21
Public Function TestPacsOrCabs(ByVal OfferNum As String, DIV As String) As Boolean
    Dim Query As String
    Dim rs As ADODB.Recordset
    
    'populate query
    Query = ReadTextFile("K:\AA\SHARE\AuditTools\rtmacros\sql\PACS_CABS_LOOKUP.sql")
    Query = Replace(Query, "(&OFFER_NUM)", OfferNum)
    Query = Replace(Query, "(&DIV)", DIV)
    
    'send query to snowflake
    Set rs = QuerySnowFlake(Query)
    
    'check if results are empty
    If IsRecordsetEmpty(rs) Then
        Exit Function
    End If
    
    'loop throguh results
    Dim returntype As String
    Do While Not rs.EOF
        returntype = rs.Fields("BILLING_SYSTEM").Value
    rs.MoveNext
    Loop
    
    'check if the return value is CABS or PACS
    If returntype = "CABS" Then
        TestPacsOrCabs = True
    Else
        TestPacsOrCabs = False
    End If
    
End Function

Public Sub INSERT_CABS_DB2_TO_WORKSHEET() 'CABS QUERY, SF QUERY OPTIONAL, DB2QUERY OPTIONAL
    'RUN QUERY
    'RUN SECOND OR THIRD QUERY
    'ADD RECORDS TO SINGLE WORKSHEET
    Dim rs As ADODB.Recordset
    Dim LastRow As Long
    Dim DB2_SQL As String
    Dim CABS_SQL As String
    Dim SF_SQL As String
    
    DB2_SQL = "Select DISTINCT RES_DIVISION, ITEM from SQLDAT3.WMALWCOM WHERE ITEM = '1010492'"
    CABS_SQL = "SELECT DISTINCT 5, AID.CORP_ITM_CD FROM DBORCABS00.ALWNC_INCOME_HDR AIH LEFT JOIN ALWNC_INCOME_DTL AID ON AID.ALWNC_INCOME_SK = AIH.ALWNC_INCOME_SK WHERE AID.CORP_ITM_CD = '48050390'"
    
    Cells(1, 1).Value = "division"
    Cells(1, 2).Value = "item"
    
    If CABS_SQL <> "" Then
        Set rs = QueryCABS(CABS_SQL)
        Range("A2").CopyFromRecordset rs
    End If
    If DB2_SQL <> "" Then
        Set rs = DB2.RunQueryRS(DB2_SQL)
        LastRow = Cells(Rows.Count, 1).End(xlUp).row + 1
        Range("A" & LastRow).CopyFromRecordset rs
    End If
    If SF_SQL <> "" Then
        Set rs = QuerySnowFlake(SF_SQL)
        LastRow = Cells(Rows.Count, 1).End(xlUp).row + 1
        Range("A" & LastRow).CopyFromRecordset rs
    End If
    
End Sub


