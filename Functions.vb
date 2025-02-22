Public Function range(start As Integer, stop1 As Variant, Optional step As Integer = 1) As Variant
    Dim result As Object
    Dim i As Integer
    
    Set result = CreateObject("Scripting.Dictionary")
    
    If IsMissing(stop1) Then
        stop1 = start
        start = 0
    End If
    
    If step > 0 Then
        For i = start To stop1 - 1 Step step
            result.Add i, True
        Next i
    Else
        i = start
        Do While i <> stop1
            result.Add i, True
            i = i + step
        Loop
    End If
    
    range = result.keys
End Function

Public Function ArrayContains(arr As Variant, value As Variant) As Boolean
    Dim element As Variant
    If IsEmpty(arr) Then
        ArrayContains = False
        Exit Function
    End If

    For Each element In arr
        If element = value Then
            ArrayContains = True
            Exit Function
        End If
    Next element
    ArrayContains = False
End Function


Public Function equalsValuesArrs(arr1 As Variant, arr2 As Variant) As Variant
    Dim value As Variant
    Dim valueToReturn As Object
    Dim i As Integer
    
    Set valueToReturn = CreateObject("Scripting.Dictionary")
    
    i = 1
    For Each value In arr1
        If ArrayContains(arr2, value) Then
            valueToReturn.Add CStr(i), value
            i = i + 1
        End If
    Next value
    
    If IsEmpty(valueToReturn.items) Then
        equalsValuesArrs = Empty
    Else
        equalsValuesArrs = valueToReturn.items
    End If
    
    Exit Function
End Function



Public Function deleteFromArr(arr As Variant, value As Variant, Optional count As Integer = -1) As Variant
    Dim newArr As Object
    Dim localCount As Integer
    Dim i As Integer
    Dim item As Variant
    
    Set newArr = CreateObject("Scripting.Dictionary")
    localCount = 1
    
    i = 0
    If count > 0 Then
        For Each item In arr
            If item <> value Then
                newArr.Add i, item
                i = i + 1
            End If
        Next item
    Else
        For Each item In arr
            If localCount <= count And item = value Then
                localCount = localCount + 1
            Else
                newArr.Add i, item
                i = i + 1
            End If
        Next item
    End If

    If newArr.count = 0 Then
        deleteFromArr = Empty
    Else
        deleteFromArr = newArr.items
    End If
End Function

Public Function addToArr(arr As Variant, value As Variant) As Variant
    Dim newArr As Object
    Dim i As Integer
    Dim item As Variant
    

    Set newArr = CreateObject("Scripting.Dictionary")

    If IsEmpty(arr) Then
        addToArr = Array(value)
        Exit Function
    End If

    i = 0
    For Each item In arr
        newArr.Add i, item
        i = i + 1
    Next item

    newArr.Add i, value
    If IsEmpty(newArr.items) Then
        addToArr = Empty
        Exit Function
    End If
    addToArr = newArr.items
End Function


