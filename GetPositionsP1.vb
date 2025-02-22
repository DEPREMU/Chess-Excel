Public Function getAvailablePosP1(piece As String) As Variant

    If playerOne(piece)("dead") Then
        getAvailablePosP1 = Empty
        Exit Function
    End If

    Dim pieceType As String
    Dim indexLetter As Integer
    Dim number As String
    Dim letter As String
    Dim btn As String
    Dim valueToReturn As Object
    Dim btnLocal As String
    Dim i As Integer
    Dim j As Integer
    Dim newNum As Integer
    Dim player As Integer
    Dim availablePos As Object
    Dim value As Variant

    Set availablePos = CreateObject("Scripting.Dictionary")
    Set valueToReturn = CreateObject("Scripting.Dictionary")

    pieceType = playerOne(piece)("type")

    Select Case pieceType

        Case "Pawn"
            letter = Mid(playerOne(piece)("newPos"), 1, 1)
            indexLetter = numbers(letter)
            number = Mid(playerOne(piece)("newPos"), 2, 1)

            ' Move forward
            If CInt(number) < 8 Then
                btn = letter & CStr(CInt(number) + 1)
                If Not buttons(btn)("isPiece") Then
                    availablePos.Add btn, True
                    If playerOne(piece)("firstMove") And Not buttons(letter & CStr(CInt(number) + 2))("isPiece") Then availablePos.Add letter & CStr(CInt(number) + 2), True
                End If
            End If

            ' Capture diagonally left
            If indexLetter > 1 And CInt(number) < 8 Then
                btn = letters(CStr(indexLetter - 1)) & CStr(CInt(number) + 1)
                If buttons(btn)("player") = 2 Then
                    availablePos.Add btn, True
                    playerTwo(buttons(btn)("piece"))("danger") = True
                    playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                End If
            End If

            ' Capture diagonally right
            If indexLetter < 8 And CInt(number) < 8 Then
                btn = letters(CStr(indexLetter + 1)) & CStr(CInt(number) + 1)
                If buttons(btn)("player") = 2 Then
                    If piece = "A1Pawn" Then MsgBox btn & " |   " & buttons(btn)("player") = 2
                    playerTwo(buttons(btn)("piece"))("danger") = True
                    playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                    availablePos.Add btn, True
                End If
            End If

        Case "Rook"
            values = getPosLeftRightTopBottomP1(piece)
            If IsEmpty(values) Then
                getAvailablePosP1 = Empty
            Else
                getAvailablePosP1 = values
            End If
            Exit Function

        Case "Knight"
            Dim newLetterIndex As Integer
            Dim offsets As Variant
            Dim position As String

            position = playerOne(piece)("newPos")
            letter = Mid(position, 1, 1)
            indexLetter = numbers(letter)
            number = CInt(Mid(position, 2, 1))

            offsets = Array(Array(2, - 1), Array(2, 1), Array( - 2, - 1), Array( - 2, 1), Array(1, - 2), Array(1, 2), Array( - 1, - 2), Array( - 1, 2))

            For i = LBound(offsets) To UBound(offsets)
                newNum = number + offsets(i)(0)
                newLetterIndex = indexLetter + offsets(i)(1)

                If newNum >= 1 And newNum <= 8 And newLetterIndex >= 1 And newLetterIndex <= 8 Then
                    btn = letters(CStr(newLetterIndex)) & CStr(newNum)
                    player = buttons(btn)("player")

                    If player = 2 Or player = 0 Then
                        If player = 2 Then
                            playerTwo(buttons(btn)("piece"))("danger") = True
                            playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                        End If
                        availablePos.Add btn, True
                    End If
                End If
            Next

        Case "Bishop"
            values = getPosBishopP1(piece)
            If IsEmpty(values) Then
                values = Empty
            Else
                For Each value In values
                    availablePos.Add value, True
                Next value
            End If

        Case "Queen"

            values = getPosLeftRightTopBottomP1(piece)
            If IsEmpty(values) Then
                values = Empty
            Else
                For Each value In values
                    availablePos.Add value, True
                Next value
            End If

            values = getPosBishopP1(piece)
            If IsEmpty(values) Then
                values = Empty
            Else
                For Each value In values
                    availablePos.Add value, True
                Next value
            End If

        Case "King"
            values = possiblePosKingP1(piece, CStr(playerOne(piece)("newPos")))
            If IsEmpty(values) Then
                getAvailablePosP1 = Empty
            Else
                getAvailablePosP1 = values
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
        getAvailablePosP1 = Empty
        Exit Function
    End If
    getAvailablePosP1 = valueToReturn.items
End Function

Public Function possiblePosKingP1(piece As String, position As String) As Variant

    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Set availablePos = CreateObject("Scripting.Dictionary")

    letter = Mid(position, 1, 1)
    indexLetter = numbers(letter)
    number = Mid(position, 2, 1)
    valuesAdded = 0

    Dim directions As Variant
    directions = Array(Array(0, 1), Array( - 1, 1), Array(1, 1), Array( - 1, 0), Array(1, 0), Array(0, - 1), Array( - 1, - 1), Array(1, - 1))

    For i = LBound(directions) To UBound(directions)

        newLetterIndex = indexLetter + directions(i)(0)
        newNumber = CInt(number) + directions(i)(1)

        If newLetterIndex >= 1 And newLetterIndex <= 8 And newNumber >= 1 And newNumber <= 8 Then
            btn = letters(CStr(newLetterIndex)) & CStr(newNumber)
            If Not buttons(btn)("isPiece") Or buttons(btn)("player") = 2 Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
        End If
    Next i

    If playerOne(piece)("firstMove") Then
        ' Enroque corto
        If Not buttons("F1")("isPiece") And Not buttons("G1")("isPiece") And playerOne("H1Rook")("firstMove") Then
            availablePos.Add "G1", True
            valuesAdded = valuesAdded + 1
        End If

        ' Enroque largo
        If Not buttons("B1")("isPiece") And Not buttons("C1")("isPiece") And Not buttons("D1")("isPiece") And playerOne("A1Rook")("firstMove") Then
            availablePos.Add "C1", True
            valuesAdded = valuesAdded + 1
        End If
    End If
    If valuesAdded = 0 Then
        possiblePosKingP1 = Empty
    Else
        possiblePosKingP1 = availablePos.keys
    End If
End Function

Public Function getPosLeftRightTopBottomP1(piece As String) As Variant
    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Set availablePos = CreateObject("Scripting.Dictionary")
    valuesAdded = 0

    letter = Mid(playerOne(piece)("newPos"), 1, 1)
    indexLetter = numbers(letter)
    number = Mid(playerOne(piece)("newPos"), 2, 1)

    ' Top
    For i = CInt(number) + 1 To 8
        btn = letter & CStr(i)
        If buttons(btn)("isPiece") Then
            If buttons(btn)("player") = 2 Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
                playerTwo(buttons(btn)("piece"))("danger") = True
                playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
            End If
            Exit For
        End If
        availablePos.Add btn, True
        valuesAdded = valuesAdded + 1
    Next i

    ' Bottom
    For Each value In range(CInt(number), 0, - 1)
        If number <> value And value Then
            btn = letter & CStr(value)
            If buttons(btn)("isPiece") Then
                If buttons(btn)("player") = 2 Then
                    availablePos.Add btn, True
                    valuesAdded = valuesAdded + 1
                    playerTwo(buttons(btn)("piece"))("danger") = True
                    playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                End If
                Exit For
            End If
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
        End If
    Next value

    ' Left
    For Each value In Array("H", "G", "F", "E", "D", "C", "B", "A")
        If indexLetter = 1 Then Exit For
        If indexLetter < numbers(value) Or value = letter Then Goto ContinueLoop

        btn = value & number
        If buttons(btn)("isPiece") Then
            If buttons(btn)("player") = 2 Then
                playerTwo(buttons(btn)("piece"))("danger") = True
                playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
            Exit For
        End If
        availablePos.Add btn, True
        valuesAdded = valuesAdded + 1

        Goto ContinueLoop
        ContinueLoop :
    Next value

    ' Right
    For Each value In Array("A", "B", "C", "D", "E", "F", "G", "H")
        If indexLetter = 8 Then Exit For
        If numbers(value) > indexLetter Then
            btn = value & number
            If buttons(btn)("isPiece") Then
                If buttons(btn)("player") = 2 Then
                    playerTwo(buttons(btn)("piece"))("danger") = True
                    playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                    availablePos.Add btn, True
                    valuesAdded = valuesAdded + 1
                End If
                Exit For
            End If
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
        End If
    Next value

    If valuesAdded = 0 Then
        getPosLeftRightTopBottomP1 = Empty
    Else
        getPosLeftRightTopBottomP1 = availablePos.keys
    End If
End Function

Public Function getPosBishopP1(piece As String) As Variant

    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Set availablePos = CreateObject("Scripting.Dictionary")
    valuesAdded = 0

    letter = Mid(playerOne(piece)("newPos"), 1, 1)
    indexLetter = numbers(letter)
    number = Mid(playerOne(piece)("newPos"), 2, 1)

    ' Top Left
    i = indexLetter
    j = CInt(number)
    Do While i > 1 And j < 8
        i = i - 1
        j = j + 1
        btn = letters(CStr(i)) & CStr(j)
        If buttons(btn)("isPiece") Then
            If buttons(btn)("player") = 2 Then
                playerTwo(buttons(btn)("piece"))("danger") = True
                playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
            Exit Do
        End If
        availablePos.Add btn, True
        valuesAdded = valuesAdded + 1
    Loop

    ' Top Right
    i = indexLetter
    j = CInt(number)
    Do While i < 8 And j < 8
        i = i + 1
        j = j + 1
        btn = letters(CStr(i)) & CStr(j)
        If buttons(btn)("isPiece") Then
            If buttons(btn)("player") = 2 Then
                playerTwo(buttons(btn)("piece"))("danger") = True
                playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
            Exit Do
        End If
        availablePos.Add btn, True
        valuesAdded = valuesAdded + 1
    Loop

    ' Bottom Left
    i = indexLetter
    j = CInt(number)
    Do While i > 1 And j > 1
        i = i - 1
        j = j - 1
        btn = letters(CStr(i)) & CStr(j)
        If buttons(btn)("isPiece") Then
            If buttons(btn)("player") = 2 Then
                playerTwo(buttons(btn)("piece"))("danger") = True
                playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
            Exit Do
        End If
        availablePos.Add btn, True
        valuesAdded = valuesAdded + 1
    Loop

    ' Bottom Right
    i = indexLetter
    j = CInt(number)
    Do While i < 8 And j > 1
        i = i + 1
        j = j - 1
        btn = letters(CStr(i)) & CStr(j)
        If buttons(btn)("isPiece") Then
            If buttons(btn)("player") = 2 Then
                playerTwo(buttons(btn)("piece"))("danger") = True
                playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
            Exit Do
        End If
        availablePos.Add btn, True
        valuesAdded = valuesAdded + 1
    Loop

    If valuesAdded = 0 Then
        getPosBishopP1 = Empty
    Else
        getPosBishopP1 = availablePos.keys
    End If
End Function