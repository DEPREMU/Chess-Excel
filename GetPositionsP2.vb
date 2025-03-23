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
    Dim i As Integer
    Dim j As Integer
    Dim newNum As Integer
    Dim player As Integer
    Dim availablePos As Object
    Dim value As Variant
    Dim posP1 As String
    Dim leftOrRight As Integer
    
    Set availablePos = CreateObject("Scripting.Dictionary")
    Set valueToReturn = CreateObject("Scripting.Dictionary")
    
    pieceType = playerTwo(piece)("type")
    
    Select Case pieceType
            
        Case "Pawn"
            letter = Mid(playerTwo(piece)("newPos"), 1, 1)
            indexLetter = numbers(letter)
            number = Mid(playerTwo(piece)("newPos"), 2, 1)
            
            For i = 0 To 1
                leftOrRight = IIf(i = 0, 1, - 1)
                If letters(CStr(indexLetter - leftOrRight)) <> "" Then
                    btn = letters(CStr(indexLetter - leftOrRight)) & number
                    If btn <> "" Then
                        If buttons(btn)("player") = 1 Then
                            If playerOne(buttons(btn)("piece"))("enPassant") Then
                                posP1 = playerOne(buttons(btn)("piece"))("newPos")
                                availablePos.Add Mid(posP1, 1, 1) & CStr(CInt(Mid(posP1, 2, 1)) - 1), True
                            End If
                        End If
                    End If
                End If
            Next i
            
            ' Move forward
            If CInt(number) > 1 Then
                btn = letter & CStr(CInt(number) - 1)
                If Not buttons(btn)("isPiece") Then
                    availablePos.Add btn, True
                    If playerTwo(piece)("firstMove") Then
                        If Not buttons(letter & CStr(CInt(number) - 2))("isPiece") Then availablePos.Add letter & CStr(CInt(number) - 2), True
                    End If
                End If
            End If
            
            ' Capture diagonally left
            If indexLetter > 1 And CInt(number) > 1 Then
                btn = letters(CStr(indexLetter - 1)) & CStr(CInt(number) - 1)
                If buttons(btn)("player") = 1 Then
                    availablePos.Add btn, True
                    playerOne(buttons(btn)("piece"))("danger") = True
                    playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
                End If
            End If
            
            ' Capture diagonally right
            If indexLetter < 8 And CInt(number) > 1 Then
                btn = letters(CStr(indexLetter + 1)) & CStr(CInt(number) - 1)
                If buttons(btn)("player") = 1 Then
                    playerOne(buttons(btn)("piece"))("danger") = True
                    playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
                    availablePos.Add btn, True
                End If
            End If
            
        Case "Rook"
            getAvailablePosP2 = getNextPosRook(piece, False, emulatePiece)
            Exit Function
            
        Case "Knight"
            getAvailablePosP2 = getNextPosKnight(piece, False)
            Exit Function

        Case "Bishop"
            getAvailablePosP2 = getNextPosBishop(piece, False, emulatePiece)
            Exit Function
            
        Case "Queen"
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
            values = possiblePosKingP2(piece, CStr(playerTwo(piece)("newPos")))
            If Not IsEmpty(values) Then
                getAvailablePosP2 = values
            Else
                getAvailablePosP2 = Empty
            End If
            Exit Function
            
    End Select
    
    i = 1
    For Each value In availablePos.keys
        If availablePos(value) Then
            valueToReturn.Add CStr(i), value
            i = i + 1
        End If
    Next value
    If i = 1 Then
        getAvailablePosP2 = Empty
        Exit Function
    End If
    getAvailablePosP2 = valueToReturn.items
End Function

Public Function possiblePosKingP2(piece As String, position As String) As Variant
    
    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Set availablePos = CreateObject("Scripting.Dictionary")
    valuesAdded = 0
    
    letter = Mid(playerTwo(piece)("newPos"), 1, 1)
    indexLetter = numbers(letter)
    number = Mid(playerTwo(piece)("newPos"), 2, 1)
    
    Dim directions As Variant
    directions = Array(Array(0, 1), Array( - 1, 1), Array(1, 1), Array( - 1, 0), Array(1, 0), Array(0, - 1), Array( - 1, - 1), Array(1, - 1))
    
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
        ' Enroque corto
        If Not buttons("F8")("isPiece") And Not buttons("G8")("isPiece") And playerTwo("H8Rook")("firstMove") And Not playerTwo("H8Rook")("dead") Then
            availablePos.Add "G8", True
            valuesAdded = valuesAdded + 1
        End If
        
        ' Enroque largo
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
