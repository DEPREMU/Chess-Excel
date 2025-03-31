Public Function range(start As Integer, stop1 As Variant, Optional Step As Integer = 1) As Variant
    Dim result As Object
    Dim i As Integer
    
    Set result = CreateObject("Scripting.Dictionary")
    
    If IsMissing(stop1) Then
        stop1 = start
        start = 0
    End If
    
    If Step > 0 Then
        For i = start To stop1 - 1 Step Step
            result.Add i, True
        Next i
    Else
        i = start
        Do While i <> stop1
            result.Add i, True
            i = i + Step
        Loop
    End If
    
    range = result.keys
    Set result = Nothing
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
    
    If IsEmpty(arr1) Or IsEmpty(arr2) Then Exit Function
    
    Set valueToReturn = CreateObject("Scripting.Dictionary")
    i = 1
    For Each value In arr1
        If ArrayContains(arr2, value) Then
            valueToReturn.Add CStr(i), value
            i = i + 1
        End If
    Next value
    
    If i = 1 Then
        Set valueToReturn = Nothing ' Limpieza antes de salir si no se encontró nada
        Exit Function
    End If
    
    equalsValuesArrs = valueToReturn.items
    Set valueToReturn = Nothing ' Limpieza de objeto para liberar memoria
End Function


Public Function deleteFromArr(arr As Variant, value As Variant, Optional count As Integer = - 1) As Variant
    Dim newArr As Object
    Dim localCount As Integer
    Dim i As Integer
    Dim item As Variant
    
    Set newArr = CreateObject("Scripting.Dictionary")
    localCount = 1
    
    If IsEmpty(arr) Then
        deleteFromArr = Empty
        Set newArr = Nothing
        Exit Function
    End If
    
    i = 0
    If count <= 0 Then
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
    
    Set newArr = Nothing
End Function


Public Function addToArr(arr As Variant, value As Variant) As Variant
    Dim newArr As Object
    Dim i As Integer
    Dim item As Variant
    
    Set newArr = CreateObject("Scripting.Dictionary")

    If IsEmpty(arr) Then
        addToArr = Array(value)
        Set newArr = Nothing
        Exit Function
    End If

    i = 1
    For Each item In arr
        If Not ArrayContains(newArr.items, item) Then
            newArr.Add i, item
            i = i + 1
        End If
    Next item

    If Not ArrayContains(newArr.items, value) Then
        newArr.Add i, value
        i = i + 1
    End If
    
    If i = 1 Then
        addToArr = Array(value)
        Set newArr = Nothing
        Exit Function
    End If
    
    addToArr = newArr.items
    Set newArr = Nothing
End Function

Public Function mergeArrs(arr1 As Variant, arr2 As Variant) As Variant
    Dim mergedArr As Object
    Dim item As Variant
    Dim i As Integer
    
    Set mergedArr = CreateObject("Scripting.Dictionary")

    If Not IsEmpty(arr1) Then
        i = 1
        For Each item In arr1
            If Not ArrayContains(mergedArr.items, item) Then
                mergedArr.Add i, item
                i = i + 1
            End If
        Next item
    End If
    
    If Not IsEmpty(arr2) Then
        For Each item In arr2
            If Not ArrayContains(mergedArr.items, item) Then
                mergedArr.Add i, item
                i = i + 1
            End If
        Next item
    End If
    
    If mergedArr.count = 0 Then
        mergeArrs = Empty
    Else
        mergeArrs = mergedArr.items
    End If
    
    Set mergedArr = Nothing
End Function