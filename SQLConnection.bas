Attribute VB_Name = "SQLConnection"
Option Explicit
Option Compare Text
Option Private Module

Const ConstStrAccess As String = "CONNECTION INPUT NEEDED"
Const ConstStrDB2P As String = "CONNECTION INPUT NEEDED"
Private Const ConstStrTeraData As String = "CONNECTION INPUT NEEDED"

Public Enum databaseLocal
    CORRESPONDENCE
    Preaudit
End Enum

Private Function getLocalDatabseString(DB As databaseLocal) As String
    Select Case DB
        Case CORRESPONDENCE: getLocalDatabseString = "CONNECTION INPUT NEEDED"
        Case Preaudit: getLocalDatabseString = "CONNECTION INPUT NEEDED"
    End Select
End Function

Public Function QueryTeraData(ByVal SQL As String) As Variant
    Static Username As String
    Static Password As String
    Static error As Boolean
    
    If Username = "" Or Password = "" Or error Then
Application.StatusBar = "Login to continue"
        With New LoginForm
            .Caption = "Login to Teradata"
            .Username = Username
            .Password = Password
            .show
            If .IsCancelled Then
                Exit Function
            Else
                Username = .Username
                Password = .Password
                error = False
            End If
        End With
    End If
    
Application.StatusBar = "Creating Connection String"
    Dim connectionString As String
    connectionString = StringInterpolation("CONNECTION INPUT NEEDED", Username, Password)
    
    'OPEN CONNECTION TO DATABASE
    Dim conn As ADODB.connection
    Set conn = New ADODB.connection
    conn.connectionString = connectionString
    
    On Error GoTo CheckPassword
Application.StatusBar = "Attempting to connect to database"
    conn.Open
    'On Error GoTo CloseConnection
    
Application.StatusBar = "Running Query..."
    'OPEN DATA AND GET RECORDSET
    Dim rs As New ADODB.Recordset
    With rs
        .ActiveConnection = conn
        .Source = SQL
        .LockType = adLockReadOnly      'MAKES CONNECTION READ ONLY
        .CursorType = adOpenForwardOnly 'FREQUENCY OF CHECKING DATABASE - SET TO ONCE
        .Open
    End With
    
    'On Error GoTo CloseRecordset
    QueryTeraData = ArrayFromRecordset(rs, True)
    
    'ERROR HANDLING
    On Error GoTo 0
CloseRecordset:
    rs.Close

CloseConnection:
    conn.Close
    
Application.StatusBar = ""
    Exit Function
CheckPassword:
    
    If err.Number = -2147467259 Then
        MsgBox err.Description, vbCritical
        error = True
        
        'QueryDb2 (sql)
    End If
    
Application.StatusBar = ""
End Function

Public Function UpdateDb2(ByVal SQL As String) As Variant
    Static Username As String
    Static Password As String
    Static error As Boolean
    
    If Username = "" Or Password = "" Or error Then
        With New LoginForm
            .Caption = "Login to DB2"
            .Username = Username
            .Password = Password
            .show
            If .IsCancelled Then
                Exit Function
            Else
                Username = .Username
                Password = .Password
                error = False
            End If
        End With
    End If
    
    Dim connectionString As String
    connectionString = StringInterpolation("Provider=MSDASQL.1;Persist Security Info=True;Extended Properties=DSN=DB2P;UID=${0};PWD=${1};", Username, Password)
    
    'OPEN CONNECTION TO DATABASE
    Dim conn As ADODB.connection
    Set conn = New ADODB.connection
    conn.connectionString = connectionString
    
    On Error GoTo CheckPassword
    conn.Open
    'On Error GoTo CloseConnection
    
    'OPEN DATA AND GET RECORDSET
    Dim rs As New ADODB.Recordset
    With rs
        .ActiveConnection = conn
        .Source = SQL
        .LockType = adLockOptimistic      'MAKES CONNECTION READ ONLY
        .CursorType = adOpenForwardOnly 'FREQUENCY OF CHECKING DATABASE - SET TO ONCE
        .Open
    End With
    
    'ERROR HANDLING
    Exit Function
CheckPassword:
    
    If err.Number = -2147467259 Then
        MsgBox err.Description, vbCritical
        error = True
        'QueryDb2 (sql)
    End If
End Function

'==============================================================================
' MAIN SUB. CREATES CONNECTION TO DB, LOOPS RECORDS & MACTCHS INVOICE #'S,
' APPIES UPDATE TO PAYBACK DATE. ANY ERRORS CANCELS THE COMMIT TO DB.
' -----------------------------------------------------------------------------
' PACKAGE INCLUDES: ArrayValueExists
'==============================================================================
Public Function sqlDatabase(dbString As databaseLocal, SQL As String) As Variant
    'DECLARE VARIABLES
    Dim temp As Variant
    Dim cell As Range
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim myConn As ADODB.connection
    Dim myData As ADODB.Recordset
    Dim myField As ADODB.field 'HEADINGS
    
    'INITIAL SET
    Set myData = New ADODB.Recordset
    Set myConn = New ADODB.connection
    myConn.connectionString = getLocalDatabseString(dbString)
    
    'OPEN CONNECTION TO DATABASE
    myConn.Open
    On Error GoTo CloseConnection

    With myData
        .ActiveConnection = myConn
        .Source = SQL
        .LockType = adLockReadOnly      'MAKES CONNECTION READ ONLY
        .CursorType = adOpenForwardOnly 'FREQUENCY OF CHECKING DATABASE - SET TO ONCE
        .Open
    End With
    
    'SETTING RECORDS TO A NEW WORKSHEET
    On Error GoTo CloseRecordset
    temp = ArrayFromRecordset(myData)
    sqlDatabase = temp

    'ERROR HANDLING
    On Error GoTo 0
CloseRecordset:
    myData.Close

CloseConnection:
    myConn.Close
End Function

'==============================================================================
'DEVELOPER NOTES
'CREATED: NICHOLAS ACKERMAN <nicholas.ackerman@albertson.com>
' -----------------------------------------------------------------------------
' PACKAGE INCLUDES:
'PURPOSE: SQL CONNECTION FOR SNOWFLAKE
'CURRENT FUNCTIONS:INSERT=(INSERT,DELETE,UPDATE), PULL
'==============================================================================

'Created 10/28/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
Public Function IsRecordsetEmpty(ByRef rs As ADODB.Recordset) As Boolean
    IsRecordsetEmpty = rs.EOF
End Function

'Created 10/28/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
Public Function RunQuery(connectionString As String, Query As String) As Recordset
    Set RunQuery = New ADODB.Recordset
    
    Dim connection As New ADODB.connection
    
    ' Potential Error, bad connection string, network error.
    connection.Open connectionString
    RunQuery.ActiveConnection = connection
    
    ' Petential Error, bad SQL, wrong SQL.
    'Debug.Print query
    RunQuery.Open Query
    
End Function

'Created 10/28/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
Public Function QuerySnowFlake(ByVal Query As String) As ADODB.Recordset
    Dialog.show "Running query...", "This might take a few seconds"
    Set QuerySnowFlake = RunQuery(ENV.use("SNOW_FLAKE_CONNECTION_STRING"), Query)
    Dialog.Hide
End Function

'Created 10/28/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose: To insert data into a Snowflake database
'Requires for a query to be sent to this function
Public Function InsertDataIntoSnowflake(Query As String) As Integer
On Error GoTo catch

    Dim conn As ADODB.connection
    Set conn = New ADODB.connection
    
    conn.Open ENV.use("SNOW_FLAKE_CONNECTION_STRING")
    conn.Execute Query
    conn.Close
    
    'No Error is Found
    InsertDataIntoSnowflake = 0
    
    Exit Function
catch:
    Dim ErrorMessage As String
    ErrorMessage = err.Description

    Application.ScreenUpdating = True
    Console.error ErrorMessage, "SQLConnection.InsertDataIntoSnowflake - CONTACT BA"
    
    DisplayErrorMessage ErrorMessage
    
    'Error is found
    InsertDataIntoSnowflake = 1
    
End Function

'Created 10/28/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
Public Function QuerySnowFlakeEDM(ByVal Query As String) As ADODB.Recordset
    Dialog.show "Running query...", "This might take a few seconds"
    Set QuerySnowFlakeEDM = RunQuery(ENV.use("SNOW_FLAKE_CONNECTION_STRING"), Query)
    Dialog.Hide
End Function

'Created 10/28/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
'Purpose: To insert data into a Snowflake database
'Requires for a query to be sent to this function
Public Function InsertDataIntoSnowflakeEDM(Query As String) As Integer
On Error GoTo catch

    Dim conn As ADODB.connection
    Set conn = New ADODB.connection
    
    conn.Open ENV.use("SNOW_FLAKE_CONNECTION_STRING")
    conn.Execute Query
    conn.Close
    
    'No Error is Found
    InsertDataIntoSnowflakeEDM = 0
    
    Exit Function
catch:
    Dim ErrorMessage As String
    ErrorMessage = err.Description

    Application.ScreenUpdating = True
    Console.error ErrorMessage, "SQLConnection.InsertDataIntoSnowflakeEDM - CONTACT BA"
    
    DisplayErrorMessage ErrorMessage
    
    'Error is found
    InsertDataIntoSnowflakeEDM = 1
    
End Function

Public Sub ImportSnowflakeTable(ByVal SQL As String)
    Dim cnnConnect As ADODB.connection
    Dim rstRecordset As ADODB.Recordset
    
    Set cnnConnect = New ADODB.connection
    cnnConnect.Open ENV.use("SNOW_FLAKE_CONNECTION_STRING")
    cnnConnect.CommandTimeout = 180
    
    Dialog.show "Running query...", "This might take a few seconds"
    
    Set rstRecordset = New ADODB.Recordset
    rstRecordset.Open _
        Source:=SQL, _
        ActiveConnection:=cnnConnect, _
        CursorType:=adOpenDynamic, _
        LockType:=adLockReadOnly, _
        options:=adCmdText
     
    With ActiveSheet.QueryTables.Add( _
            connection:=rstRecordset, _
            destination:=Range("A1"))
        .name = "Contact List"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = True
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    
    Dialog.Hide
End Sub
'Created 10/28/2020
'BY:@Nicholas Ackerman <nicholas.ackerman@albertsons.com>
Public Function QueryCABS(ByVal Query As String) As ADODB.Recordset
    Dialog.show "Running query...", "This might take a few seconds"
    Set QueryCABS = RunQuery(ENV.use("CABS_CONNECTION_STRING"), Query)
    Dialog.Hide
End Function

Public Sub ImportCABSTable(ByVal SQL As String, Optional StartingRng As String)
    Dim cnnConnect As ADODB.connection
    Dim rstRecordset As ADODB.Recordset
    
    Set cnnConnect = New ADODB.connection
    cnnConnect.Open ENV.use("CABS_CONNECTION_STRING")
    cnnConnect.CommandTimeout = 180
    
    Dialog.show "Running query...", "This might take a few seconds"
    
    Set rstRecordset = New ADODB.Recordset
    rstRecordset.Open _
        Source:=SQL, _
        ActiveConnection:=cnnConnect, _
        CursorType:=adOpenDynamic, _
        LockType:=adLockReadOnly, _
        options:=adCmdText
     
    With ActiveSheet.QueryTables.Add( _
            connection:=rstRecordset, _
            destination:=Range(StartingRng))
        .name = "Contact List"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = True
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    
    Dialog.Hide
End Sub

