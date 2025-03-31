Public Function getNextPosKing(piece As String, position As String, boolPlayerOne As Boolean) As Variant
    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Dim player As Object
    Dim opponent As Integer
    Dim rowShortCastle As String
    Dim rowLongCastle As String
    Dim shortRook As String
    Dim longRook As String
    Set availablePos = CreateObject("Scripting.Dictionary")
    
    ' Determinar si es jugador 1 o 2
    If boolPlayerOne Then
        Set player = playerOne
        opponent = 2
        rowShortCastle = "G1"
        rowLongCastle = "C1"
        shortRook = "H1Rook"
        longRook = "A1Rook"
    Else
        Set player = playerTwo
        opponent = 1
        rowShortCastle = "G8"
        rowLongCastle = "C8"
        shortRook = "H8Rook"
        longRook = "A8Rook"
    End If
    
    letter = Mid(player(piece)("newPos"), 1, 1)
    indexLetter = numbers(letter)
    number = Mid(player(piece)("newPos"), 2, 1)
    
    Dim directions As Variant
    directions = Array(Array(0, 1), Array( - 1, 1), Array(1, 1), Array( - 1, 0), Array(1, 0), Array(0, - 1), Array( - 1, - 1), Array(1, - 1))

    valuesAdded = 0

    For i = LBound(directions) To UBound(directions)
        newLetterIndex = indexLetter + directions(i)(0)
        newNumber = CInt(number) + directions(i)(1)

        If newLetterIndex >= 1 And newLetterIndex <= 8 And newNumber >= 1 And newNumber <= 8 Then
            btn = letters(CStr(newLetterIndex)) & CStr(newNumber)
            If Not buttons(btn)("isPiece") Or buttons(btn)("player") = opponent Then
                If btn <> player(piece)("newPos") Then
                    availablePos.Add btn, True
                    valuesAdded = valuesAdded + 1
                End If
            End If
        End If
    Next i

    ' Enroque
    If player(piece)("firstMove") Then
        ' Enroque corto
        If Not buttons(Mid(rowShortCastle, 1, 1) & CStr(Mid(rowShortCastle, 2, 1) - 1))("isPiece") And _
                 Not buttons(rowShortCastle)("isPiece") And _
                player(shortRook)("firstMove") And Not player(shortRook)("dead") Then
            availablePos.Add rowShortCastle, True
            valuesAdded = valuesAdded + 1
        End If

        ' Enroque largo
        If Not buttons("B" & Mid(rowLongCastle, 2, 1))("isPiece") And _
                 Not buttons("C" & Mid(rowLongCastle, 2, 1))("isPiece") And _
                 Not buttons("D" & Mid(rowLongCastle, 2, 1))("isPiece") And _
                player(longRook)("firstMove") And Not player(longRook)("dead") Then
            availablePos.Add rowLongCastle, True
            valuesAdded = valuesAdded + 1
        End If
    End If

    If valuesAdded = 0 Then
        getNextPosKing = Empty
    Else
        getNextPosKing = availablePos.keys
    End If
    Set availablePos = Nothing
    Set player = Nothing
    
End Function