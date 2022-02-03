Attribute VB_Name = "BluezoneFactory"
Option Explicit

'/**
' * Bluezone factory method.
' *
' * @author: Robert Todar <robert.todar@albertsons.com>
' */
Public Function TryGetActiveBlueZone(ByRef outBluezone As BlueZone) As Boolean
    Dim bz As BlueZone
    Set bz = New BlueZone
    
    Dim ConnectionError As String
    If bz.TryConnectToActiveSession(True, ConnectionError) = False Then
        MsgBox "Unable to connect to BlueZone: " & ConnectionError
        Exit Function
    End If
    
    TryGetActiveBlueZone = True
    Set outBluezone = bz
End Function
