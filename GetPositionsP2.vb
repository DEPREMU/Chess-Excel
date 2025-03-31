Public Function getAvailablePosP2(piece As String, Optional emulatePiece As Variant) As Variant
    If playerTwo(piece)("dead") Then
        getAvailablePosP2 = Empty
        Exit Function
    End If
    
    If IsMissing(emulatePiece) Then emulatePiece = Array("", "")
    
    Dim pieceType As String
    Dim indexLetter As Integer
    Dim number As String
    Dim letter As String
    Dim btn As String
    Dim valueToReturn As Object
    Dim availablePos As Object
    Dim value As Variant
    Dim posP1 As String
    Dim leftOrRight As Integer
    
    Set availablePos = CreateObject("Scripting.Dictionary")
    Set valueToReturn = CreateObject("Scripting.Dictionary")
    
    pieceType = playerTwo(piece)("type")
    
    Select Case pieceType
        Case "Pawn"
            getAvailablePosP2 = getNextPosPawn(piece, False)
            Set availablePos = Nothing
            Set valueToReturn = Nothing
            Exit Function
        Case "Rook"
            getAvailablePosP2 = getNextPosRook(piece, False, emulatePiece)
            Set availablePos = Nothing
            Set valueToReturn = Nothing
            Exit Function

        Case "Knight"
            getAvailablePosP2 = getNextPosKnight(piece, False)
            Set availablePos = Nothing
            Set valueToReturn = Nothing
            Exit Function

        Case "Bishop"
            getAvailablePosP2 = getNextPosBishop(piece, False, emulatePiece)
            Set availablePos = Nothing
            Set valueToReturn = Nothing
            Exit Function
            
        Case "Queen"
            Dim values As Variant
            values = getNextPosRook(piece, False, emulatePiece)
            If Not IsEmpty(values) Then
                For Each value In values
                    availablePos.Add value, True
                Next value
            End If

            values = getNextPosBishop(piece, False, emulatePiece)
            If Not IsEmpty(values) Then
                For Each value In values
                    availablePos.Add value, True
                Next value
            End If

        Case "King"
            getAvailablePosP2 = getNextPosKing(piece, CStr(playerTwo(piece)("newPos")), False)
            Set availablePos = Nothing
            Set valueToReturn = Nothing
            Exit Function
    End Select
    
    i = 1
    For Each value In availablePos.keys
        If availablePos(value) Then
            valueToReturn.Add CStr(i), value
            i = i + 1
        End If
    Next value
    Set availablePos = Nothing
    
    If i = 1 Then
        getAvailablePosP2 = Empty
        Set valueToReturn = Nothing
        Exit Function
    End If
    
    getAvailablePosP2 = valueToReturn.items
    Set valueToReturn = Nothing
End Function

Public Function possiblePosKingP2(piece As String, position As String) As Variant
    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Set availablePos = CreateObject("Scripting.Dictionary")
    
    letter = Mid(playerTwo(piece)("newPos"), 1, 1)
    indexLetter = numbers(letter)
    number = Mid(playerTwo(piece)("newPos"), 2, 1)
    
    Dim directions As Variant
    directions = Array(Array(0, 1), Array( - 1, 1), Array(1, 1), Array( - 1, 0), Array(1, 0), Array(0, - 1), Array( - 1, - 1), Array(1, - 1))
    
    valuesAdded = 0
    
    For i = LBound(directions) To UBound(directions)
        newLetterIndex = indexLetter + directions(i)(0)
        newNumber = CInt(number) + directions(i)(1)
        
        If newLetterIndex >= 1 And newLetterIndex <= 8 And newNumber >= 1 And newNumber <= 8 Then
            btn = letters(CStr(newLetterIndex)) & CStr(newNumber)
            If Not buttons(btn)("isPiece") Or buttons(btn)("player") = 1 And btn <> playerTwo(piece)("newPos") Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
        End If
    Next i
    
    If playerTwo(piece)("firstMove") Then
        ' Short castling
        If Not buttons("F8")("isPiece") And Not buttons("G8")("isPiece") And playerTwo("H8Rook")("firstMove") And Not playerTwo("H8Rook")("dead") Then
            availablePos.Add "G8", True
            valuesAdded = valuesAdded + 1
        End If
        
        ' Long castling
        If Not buttons("B8")("isPiece") And Not buttons("C8")("isPiece") And Not buttons("D8")("isPiece") And playerTwo("A8Rook")("firstMove") And Not playerTwo("A8Rook")("dead") Then
            availablePos.Add "C8", True
            valuesAdded = valuesAdded + 1
        End If
    End If
    
    If valuesAdded = 0 Then
        possiblePosKingP2 = Empty
    Else
        possiblePosKingP2 = availablePos.keys
    End If
End Function