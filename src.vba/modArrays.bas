Option Explicit

Function IsArrayStringInitialised(ByRef a() As String) As Boolean
    On Error Resume Next
    IsArrayStringInitialised = IsNumeric(UBound(a))
    On Error GoTo 0
End Function

Public Function IsInArray(arrayData, searchData) As Integer
    Dim found As Integer, i As Integer
    
    If IsEmpty(arrayData) Then
        found = -1
    Else
        found = -1
        For i = 0 To UBound(arrayData)
            If arrayData(i) = searchData Then
                found = i
                Exit For
            End If
        Next i
    End If
    
    IsInArray = found
End Function

Public Function ClearArray(arr) As Variant
    Dim newarr As Variant
    
    'Erase arr
    
    ClearArray = newarr
End Function

Public Function AddToArray(arr, elem) As Variant
    If (IsNull(elem) Or IsEmpty(elem)) Then
        ' nothing to do
        AddToArray = arr
        Exit Function
    Else
        
        If IsEmpty(arr) Then
            'ReDim newarr(0 To 0)
            'newarr(0) = elem
            'AddToArray = newarr
            AddToArray = Array(elem)
        Else
            ReDim newarr(0 To UBound(arr) + 1)
            Dim i As Integer
            
            If IsObject(elem) Then
                Set newarr(UBound(arr) + 1) = elem
            Else
                newarr(UBound(arr) + 1) = elem
            End If
            
            For i = 0 To UBound(arr)
            
                If IsObject(arr(i)) Then
                    Set newarr(i) = arr(i)
                Else
                    newarr(i) = arr(i)
                End If
                
            Next i
            
            AddToArray = newarr
            Exit Function
        End If
    End If
End Function

Public Function RemoveFromArray(arr, elem, Optional iPos As Integer = -1) As Variant
    Dim newarr As Variant, bFound As Boolean
    
    newarr = arr
    bFound = True
    
    If iPos = -1 Then
        Do Until Not bFound
            newarr = RemoveFromArrayFirstOccur(newarr, elem, bFound, iPos)
        Loop
    Else
        newarr = RemoveFromArrayFirstOccur(newarr, elem, bFound, iPos)
    End If
    
    RemoveFromArray = newarr
End Function

Private Function RemoveFromArrayFirstOccur(arr, elem, ByRef bFound As Boolean, Optional iPos As Integer = -1) As Variant
    
    bFound = False
    If IsEmpty(arr) Then
        RemoveFromArrayFirstOccur = arr
    Else
        Dim newarr As Variant
        Dim i As Integer, iPosfound As Integer
            
        If UBound(arr) = 0 Then
            newarr = Empty
        Else
            ReDim newarr(0 To UBound(arr) - 1)
        End If
            
        iPosfound = iPos
        For i = 0 To UBound(arr)
            If iPosfound = -1 Then
                If arr(i) = elem Then
                    iPosfound = i
                End If
            End If
                
            If iPosfound = -1 Then
                ' not found yet
                If i < UBound(arr) Then
                    newarr(i) = arr(i)
                End If
            ElseIf i = iPosfound Then
                ' found at the current position
                bFound = True
            ElseIf i < iPosfound Then
                ' found at not reached yet position
                newarr(i) = arr(i)
            Else
                ' found at already reached position
                newarr(i - 1) = arr(i)
            End If
            
        Next i
        
        If bFound Then
            RemoveFromArrayFirstOccur = newarr
        Else
            RemoveFromArrayFirstOccur = arr
        End If
        
    End If
End Function

Public Function AddToSplitStr(sSplitted As String, sDelim As String, sElem As String) As String
    If sSplitted = "" Then
        AddToSplitStr = sElem
    Else
        Dim arr As Variant, newarr As Variant
        arr = Split(sSplitted, sDelim)
        newarr = AddToArray(arr, sElem)
        
        AddToSplitStr = Join(newarr, sDelim)
    End If
    
End Function

Public Function RemoveFromSplitStr(sSplitted As String, sDelim As String, sElem As String, Optional iPos As Integer = -1) As String
    If sSplitted = "" Then
        RemoveFromSplitStr = sSplitted
    Else
        Dim arr As Variant, newarr As Variant
        arr = Split(sSplitted, sDelim)
        newarr = RemoveFromArray(arr, sElem, iPos)
        
        If IsEmpty(newarr) Then
            RemoveFromSplitStr = ""
        Else
            RemoveFromSplitStr = Join(newarr, sDelim)
        End If
    End If
    
End Function

Public Function AddToStringArray(arr() As String, elem As String, Optional posi As Integer = -1) As String()
    Dim i As Integer
    
    If Len(Join(arr)) = 0 Then
        ReDim newarr(0 To 0) As String
        
        If posi > 0 Then
            ReDim newarr(0 To posi) As String
            newarr(posi) = elem
        Else
            newarr(0) = elem
        End If
        
        AddToStringArray = newarr
        Exit Function
    Else
        ReDim newarr(0 To UBound(arr)) As String
        
        If posi = -1 Then
            ReDim newarr(0 To UBound(arr) + 1) As String
            newarr(UBound(newarr)) = elem
        Else
            If posi > UBound(arr) Then
                ReDim newarr(0 To posi) As String
                newarr(posi) = elem
            Else
                newarr(posi) = elem
            End If
        End If
        
        For i = 0 To UBound(arr)
            newarr(i) = arr(i)
        Next i
           
        AddToStringArray = newarr
        Exit Function
    End If
End Function