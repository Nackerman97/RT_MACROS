Attribute VB_Name = "Utilities"
Option Explicit
Option Compare Text

'/**
' * These are meant to be general utilities used for various vba tasks
' *
' * @author Robert Todar <robert@roberttodar.com>
' */

' |---------------|----------------------------------------|
' | Function Name |                Purpose                 |
' |---------------|----------------------------------------|
' | Contains      | See if list or string contains a value |
' | Length        | See count of items or string legnth    |
' | Concat        | Joins two of the same type together    |
' |---------------|----------------------------------------|

'/**
' * Run tests for this module.
' */
Private Sub testUtilityMethods()
    Dim arr As Variant
    arr = Array(1, 2, "Robert")
    
    Dim dict As New Scripting.Dictionary
    dict.Add "name", "Robert"
    
    Dim col As New Collection
    col.Add "Robert", "name"
    
    Dim Rng As Range
    Set Rng = Range("A2:A4")
    
    Dim Str As String
    Str = "Hi Robert"
    
    Debug.Print Length(arr)
    Debug.Print Length(dict)
    Debug.Print Length(col)
    Debug.Print Length(Rng)
    Debug.Print Length(Str)
    
    Debug.Print Contains(arr, "Robert")
    Debug.Print Contains(dict, "Robert")
    Debug.Print Contains(col, "Robert")
    Debug.Print Contains(Rng, "Robert")
    Debug.Print Contains(Str, "Robert")
End Sub

Private Function testConcatFunctions()
    Dim arr1 As Variant
    arr1 = Array(1, 2, 3)
    
    Dim arr2 As Variant
    arr2 = Array(4, 5, 6, 7, 8)
    
    Dim arr3 As Variant
    arr3 = Concat(arr1, arr2)
    Debug.Print ToString(arr3)
    
    Dim col1 As New Collection
    col1.Add 1
    
    Dim col2 As New Collection
    col2.Add 2
    col2.Add 3
    
    Dim col3 As Collection
    Set col3 = Concat(col1, col2)
    
    Debug.Print ToString(col3)
    
    Dim dict1 As New Scripting.Dictionary
    dict1.Add "name", "Robert"
    
    Dim dict2 As New Scripting.Dictionary
    dict2.Add "age", 31
    
    Dim dict3 As Scripting.Dictionary
    Set dict3 = Concat(dict1, dict2)
    
    Debug.Print ToString(dict3)
End Function

'/**
' * Checks to see if some value is in some type of list. IE: Array, Collection, Dictionary.
' */
Public Function Contains(ByVal Source As Variant, ByVal Value As Variant) As Boolean
    Select Case TypeName(Source)
        Case "Variant()"
            Contains = arrayContains(Source, Value)
        
        Case "Collection"
            Contains = collectionContains(Source, Value)
            
        Case "Dictionary"
            Contains = dictionaryContains(Source, Value)
            
        Case "Range"
            Contains = rangeContains(Source, Value)
            
        Case "String"
            Contains = (InStr(Source, Value) > 0)
            
        Case Else
            err.Raise 13, "Contains", "Unknown way seeing if " & TypeName(Source) & " contains " & Value
    End Select
End Function

'/**
' * Checks to see if some value is in an Range.
' */
Private Function rangeContains(ByVal Source As Range, ByVal Value As Variant) As Boolean
    Dim cell As Range
    For Each cell In Source
        If cell.Value = Value Then
            rangeContains = True
            Exit Function
        End If
    Next cell
End Function

'/**
' * Checks to see if some value is in an array.
' */
Private Function arrayContains(ByVal Source As Variant, ByVal Value As Variant) As Boolean
    Dim index As Long
    For index = LBound(Source) To UBound(Source)
        If Source(index) = Value Then
            arrayContains = True
            Exit Function
        End If
    Next index
End Function

'/**
' * Checks to see if some value is in a collection.
' */
Private Function collectionContains(ByVal Source As Collection, ByVal Value As Variant) As Boolean
    Dim index As Long
    For index = 1 To Source.Count
        If Source.item(index) = Value Then
            collectionContains = True
            Exit Function
        End If
    Next index
End Function

'/**
' * Checks to see if some value is in a dictionary.
' */
Private Function dictionaryContains(ByVal Source As Scripting.Dictionary, ByVal Value As Variant) As Boolean
    Dim index As Long
    For index = 0 To Source.Count - 1
        If Source.items(index) = Value Then
            dictionaryContains = True
            Exit Function
        End If
    Next index
End Function

'/**
' * Get's the Length (count) of items in some type of list.
' */
Public Function Length(ByVal Source As Variant) As Long
    Select Case TypeName(Source)
        Case "Variant()"
            Length = UBound(Source) - LBound(Source) + 1
            
        Case "Dictionary", "Collection"
            Length = Source.Count
            
        Case "Range"
            Length = Source.Cells.Count
            
        Case "String"
            Length = Len(Source)
            
        Case Else
            err.Raise 13, "Length", "Unknown Length of " & TypeName(Source)
    End Select
End Function

'/**
' * Joins two of the same type of list.
' */
Public Function Concat(ByVal sourceOne As Variant, ByVal sourceTwo As Variant) As Variant
    If TypeName(sourceOne) <> TypeName(sourceTwo) Then
        err.Raise 13, "Concat", TypeName(sourceOne) & " does not match " & TypeName(sourceTwo)
    End If
    
    Select Case TypeName(sourceOne)
        Case "Variant()"
            Concat = arrayConcat(sourceOne, sourceTwo)
            
        Case "Collection"
            Set Concat = collectionConcat(sourceOne, sourceTwo)
        
        Case "Dictionary"
            Set Concat = dictionaryConcat(sourceOne, sourceTwo)
            
        Case "Range"
            Set Concat = Union(sourceOne, sourceTwo)
            
        Case "String"
            Concat = sourceOne & sourceTwo
            
        Case Else
            err.Raise 13, "Concat", "Unknown way to concat " & TypeName(sourceOne)
    End Select
End Function

'/**
' * Joins two arrays together.
' */
Private Function arrayConcat(ByVal sourceOne As Variant, ByVal sourceTwo As Variant) As Variant
    ' Resize temp to have room for both arrays.
    Dim temp As Variant
    ReDim temp(LBound(sourceOne) To UBound(sourceOne) + Length(sourceTwo))
    
    ' Capture length to later remove from index for source two.
    ' This is a way of resetting the index to start at the first position to sourceTwo.
    Dim sourceOneLength As Long
    sourceOneLength = Length(sourceOne)
    
    Dim index As Long
    For index = LBound(temp) To UBound(temp)
        If index > UBound(sourceOne) Then
            ' Add second array after the first is populated
            temp(index) = sourceTwo(index - sourceOneLength)
        Else
            ' Add first array
            temp(index) = sourceOne(index)
        End If
    Next index
    arrayConcat = temp
End Function

'/**
' * Joins two collections together.
' */
Private Function collectionConcat(ByVal sourceOne As Collection, ByVal sourceTwo As Collection) As Collection
    Set collectionConcat = New Collection
    Dim index As Long
    
    For index = 1 To sourceOne.Count
        collectionConcat.Add sourceOne.item(index)
    Next index
    
    For index = 1 To sourceTwo.Count
        collectionConcat.Add sourceTwo.item(index)
    Next index
End Function

'/**
' * Joins two dictionaries together.
' */
Private Function dictionaryConcat(ByVal sourceOne As Scripting.Dictionary, ByVal sourceTwo As Scripting.Dictionary) As Scripting.Dictionary
    Set dictionaryConcat = New Scripting.Dictionary
    Dim index As Long
    
    For index = 0 To sourceOne.Count - 1
        dictionaryConcat.Add sourceOne.Keys(index), sourceOne.items(index)
    Next index
    
    For index = 0 To sourceTwo.Count - 1
        dictionaryConcat.Add sourceTwo.Keys(index), sourceTwo.items(index)
    Next index
End Function

'/**
' * Removes duplicates from some type of list.
' */
Private Function RemoveDuplicates(ByVal Source As Variant) As Variant
    ' TODO: Need to write this formula and then make public.
End Function

