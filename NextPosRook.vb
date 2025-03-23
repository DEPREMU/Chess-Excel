Option Explicit

Public Function getNextPosRook(piece As String, boolPlayerOne as boolean, Optional emulatePiece As Variant) As Variant
    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Dim currentPos As String
    Dim lastPos As Variant
    Dim lastPiece As Variant
    Dim lastPlayer As Integer
    Dim lastIsPiece As Boolean
    Dim lastIsPieceBtns As Boolean
    dim i as integer
    dim btn as string
    Set availablePos = CreateObject("Scripting.Dictionary")
    valuesAdded = 0
    
    If IsMissing(emulatePiece) Then emulatePiece = Array("", "")
    If emulatePiece(0) <> "" And emulatePiece(1) <> "" Then
        lastPos = getPosPlayer(CStr(emulatePiece(0)), Not boolPlayerOne)
        lastPlayer = buttons(lastPos)("player")
        lastPiece = buttons(lastPos)("piece")
        lastIsPiece = buttons(lastPos)("isPiece")
        lastIsPieceBtns = buttons(emulatePiece(1))("isPiece")

        buttons(lastPos)("isPiece") = False
        buttons(lastPos)("player") = 0
        If boolPlayerOne Then
            playerTwo(emulatePiece(0))("newPos") = emulatePiece(1)
        Else
            playerOne(emulatePiece(0))("newPos") = emulatePiece(1)
        End If
        buttons(emulatePiece(1))("isPiece") = True
    End If
    
    currentPos = getPosPlayer(piece, boolPlayerOne)
    
    letter = Mid(currentPos, 1, 1)
    indexLetter = numbers(letter)
    number = Mid(currentPos, 2, 1)
    
    ' Top
    For i = CInt(number) + 1 To 8
        btn = letter & CStr(i)
        If emulatePiece(1) = btn Then
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
            If boolPlayerOne Then
                playerTwo(emulatePiece(0))("danger") = True
                playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Else
                playerOne(emulatePiece(0))("danger") = True
                playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
            End If
            Exit For
        ElseIf buttons(btn)("isPiece") Then
            If (buttons(btn)("player") = 2 And boolPlayerOne) Or _
                    (buttons(btn)("player") = 1 And Not boolPlayerOne) Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
                If boolPlayerOne Then
                    playerTwo(buttons(btn)("piece"))("danger") = True
                    playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                Else
                    playerOne(buttons(btn)("piece"))("danger") = True
                    playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
                End If
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
            If emulatePiece(1) = btn Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
                If boolPlayerOne Then
                    playerTwo(emulatePiece(0))("danger") = True
                    playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
                Else
                    playerOne(emulatePiece(0))("danger") = True
                    playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
                End If
                Exit For
            ElseIf buttons(btn)("isPiece") Then
                If (buttons(btn)("player") = 2 And boolPlayerOne) Or _
                        (buttons(btn)("player") = 1 And Not boolPlayerOne) Then

                    availablePos.Add btn, True
                    valuesAdded = valuesAdded + 1
                    If boolPlayerOne Then
                        playerTwo(buttons(btn)("piece"))("danger") = True
                        playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                    Else
                        playerOne(buttons(btn)("piece"))("danger") = True
                        playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
                    End If
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
        If emulatePiece(1) = btn Then
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
            If boolPlayerOne Then
                playerTwo(emulatePiece(0))("danger") = True
                playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Else
                playerOne(emulatePiece(0))("danger") = True
                playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
            End If
            Exit For
        ElseIf buttons(btn)("isPiece") Then
            If (buttons(btn)("player") = 2 And boolPlayerOne) Or _
                    (buttons(btn)("player") = 1 And Not boolPlayerOne) Then
                If boolPlayerOne Then
                    playerTwo(buttons(btn)("piece"))("danger") = True
                    playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                Else
                    playerOne(buttons(btn)("piece"))("danger") = True
                    playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
                End If
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
            If emulatePiece(1) = btn Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
                If boolPlayerOne Then
                    playerTwo(emulatePiece(0))("danger") = True
                    playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
                Else
                    playerOne(emulatePiece(0))("danger") = True
                    playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
                End If
                Exit For
            ElseIf buttons(btn)("isPiece") Then
                If (buttons(btn)("player") = 2 And boolPlayerOne) Or _
                        (buttons(btn)("player") = 1 And Not boolPlayerOne) Then
                    If boolPlayerOne Then
                        playerTwo(buttons(btn)("piece"))("danger") = True
                        playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                    Else
                        playerOne(buttons(btn)("piece"))("danger") = True
                        playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
                    End If
                    availablePos.Add btn, True
                    valuesAdded = valuesAdded + 1
                End If
                Exit For
            End If
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
        End If
    Next value
    
    If emulatePiece(0) <> "" And emulatePiece(1) <> "" Then
        If boolPlayerOne Then
            playerTwo(emulatePiece(0))("newPos") = lastPos
        Else
            playerOne(emulatePiece(0))("newPos") = lastPos
        End If
        lastPos = getPosPlayer(CStr(emulatePiece(0)), Not boolPlayerOne)
        buttons(lastPos)("isPiece") = lastIsPiece
        buttons(lastPos)("player") = lastPlayer
        buttons(lastPos)("piece") = lastPiece
        buttons(emulatePiece(1))("isPiece") = lastIsPieceBtns
    End If

    If valuesAdded = 0 Then
        getPosRook = Empty
    Else
        getPosRook = availablePos.keys
    End If
End Function