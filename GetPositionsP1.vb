Public Function getAvailablePosP1(piece As String, Optional emulatePiece As Variant) As Variant
    getAvailablePosP1 = Empty
    
    If playerOne(piece)("dead") Then Exit Function
    
    If IsMissing(emulatePiece) Then emulatePiece = Array("", "")
    
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
                    If playerOne(piece)("firstMove") Then
                        If Not buttons(letter & CStr(CInt(number) + 2))("isPiece") Then availablePos.Add letter & CStr(CInt(number) + 2), True
                    End If
                End If
            End If
            
            ' Capture diagonally left
            If indexLetter > 1 And CInt(number) < 8 Then
                btn = letters(CStr(indexLetter - 1)) & CStr(CInt(number) + 1)
                If buttons(btn)("player") = 2 Then
                    If playerTwo(buttons(btn)("piece"))("enPassant") Then
                        btn = Mid(buttons(btn)("piece"), 1, 1) & CStr(CInt(Mid(buttons(btn)("piece"), 2, 1)) - 1)
                        If Not buttons(btn)("isPiece") Then availablePos.Add btn, True
                    Else
                        availablePos.Add btn, True
                        playerTwo(buttons(btn)("piece"))("danger") = True
                        playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                    End If
                End If
            End If
            
            ' Capture diagonally right
            If indexLetter < 8 And CInt(number) < 8 Then
                btn = letters(CStr(indexLetter + 1)) & CStr(CInt(number) + 1)
                If buttons(btn)("player") = 2 Then
                    If playerTwo(buttons(btn)("piece"))("enPassant") Then
                        btn = Mid(buttons(btn)("piece"), 1, 1) & CStr(CInt(Mid(buttons(btn)("piece"), 2, 1)) - 1)
                        If Not buttons(btn)("isPiece") Then availablePos.Add btn, True
                    Else
                        playerTwo(buttons(btn)("piece"))("danger") = True
                        playerTwo(buttons(btn)("piece"))("piecesEater") = addToArr(playerTwo(buttons(btn)("piece"))("piecesEater"), piece)
                        availablePos.Add btn, True
                    End If
                End If
            End If
            
        Case "Rook"
            values = getPosRookP1(piece, emulatePiece)
            If Not IsEmpty(values) Then getAvailablePosP1 = values
            Exit Function
            
        Case "Knight"
            Dim newLetterIndex As Integer
            Dim offsets As Variant
            Dim position As String
            
            position = playerOne(piece)("newPos")
            letter = Mid(position, 1, 1)
            indexLetter = numbers(letter)
            number = CInt(Mid(position, 2, 1))
            
            offsets = Array(Array(2, -1), Array(2, 1), Array(-2, -1), Array(-2, 1), Array(1, -2), Array(1, 2), Array(-1, -2), Array(-1, 2))
            
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
            If Not IsEmpty(values) Then getAvailablePosP1 = values
            Exit Function
            
        Case "Queen"
            
            values = getPosRookP1(piece, emulatePiece)
            If Not IsEmpty(values) Then
                For Each value In values
                    availablePos.Add value, True
                Next value
            End If
            
            values = getPosBishopP1(piece, emulatePiece)
            If Not IsEmpty(values) Then
                For Each value In values
                    availablePos.Add value, True
                Next value
            End If
            
        Case "King"
            values = getPosKingP1(piece, CStr(playerOne(piece)("newPos")))
            If Not IsEmpty(values) Then getAvailablePosP1 = values
            Exit Function
    End Select
    
    i = 1
    For Each value In availablePos.keys
        If availablePos(value) Then
            valueToReturn.Add CStr(i), value
            i = i + 1
        End If
    Next value
    If i = 1 Then Exit Function
    
    getAvailablePosP1 = valueToReturn.items
End Function

Public Function getPosKingP1(piece As String, position As String) As Variant
    
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
    directions = Array(Array(0, 1), Array(-1, 1), Array(1, 1), Array(-1, 0), Array(1, 0), Array(0, -1), Array(-1, -1), Array(1, -1))
    
    For i = LBound(directions) To UBound(directions)
        
        newLetterIndex = indexLetter + directions(i)(0)
        newNumber = CInt(number) + directions(i)(1)
        
        If newLetterIndex >= 1 And newLetterIndex <= 8 And newNumber >= 1 And newNumber <= 8 Then
            btn = letters(CStr(newLetterIndex)) & CStr(newNumber)
            If Not buttons(btn)("isPiece") Or buttons(btn)("player") = 2 And btn <> playerOne(piece)("newPos") Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
        End If
    Next i
    
    If playerOne(piece)("firstMove") Then
        ' Enroque corto
        If Not buttons("F1")("isPiece") And Not buttons("G1")("isPiece") And playerOne("H1Rook")("firstMove") And Not playerOne("H1Rook")("dead") Then
            availablePos.Add "G1", True
            valuesAdded = valuesAdded + 1
        End If
        
        ' Enroque largo
        If Not buttons("B1")("isPiece") And Not buttons("C1")("isPiece") And Not buttons("D1")("isPiece") And playerOne("A1Rook")("firstMove") And Not playerOne("A1Rook")("dead") Then
            availablePos.Add "C1", True
            valuesAdded = valuesAdded + 1
        End If
    End If
    If valuesAdded = 0 Then
        getPosKingP1 = Empty
    Else
        getPosKingP1 = availablePos.keys
    End If
End Function

Public Function getPosRookP1(piece As String, Optional emulatePiece As Variant) As Variant
    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Dim lastPos As Variant
    Dim lastPiece As Variant
    Dim lastPlayer As Integer
    Dim lastIsPiece As Boolean
    Dim lastIsPieceBtns As Boolean
    Set availablePos = CreateObject("Scripting.Dictionary")
    valuesAdded = 0
    
    If IsMissing(emulatePiece) Then emulatePiece = Array("", "")
    If emulatePiece(0) <> "" And emulatePiece(1) <> "" Then
        lastPos = playerTwo(emulatePiece(0))("newPos")
        lastPlayer = buttons(playerTwo(emulatePiece(0))("newPos"))("player")
        lastPiece = buttons(playerTwo(emulatePiece(0))("newPos"))("piece")
        lastIsPiece = buttons(playerTwo(emulatePiece(0))("newPos"))("isPiece")
        lastIsPieceBtns = buttons(emulatePiece(1))("isPiece")

        buttons(playerTwo(emulatePiece(0))("newPos"))("isPiece") = False
        buttons(playerTwo(emulatePiece(0))("newPos"))("player") = 0
        playerTwo(emulatePiece(0))("newPos") = emulatePiece(1)
        buttons(emulatePiece(1))("isPiece") = True
    End If
    
    letter = Mid(playerOne(piece)("newPos"), 1, 1)
    indexLetter = numbers(letter)
    number = Mid(playerOne(piece)("newPos"), 2, 1)
    
    ' Top
    For i = CInt(number) + 1 To 8
        btn = letter & CStr(i)
        If emulatePiece(1) = btn Then
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
            playerTwo(emulatePiece(0))("danger") = True
            playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Exit For
        ElseIf buttons(btn)("isPiece") Then
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
    For Each value In range(CInt(number), 0, -1)
        If number <> value And value Then
            btn = letter & CStr(value)
            If emulatePiece(1) = btn Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
                playerTwo(emulatePiece(0))("danger") = True
                playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
                Exit For
            ElseIf buttons(btn)("isPiece") Then
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
        If indexLetter < numbers(value) Or value = letter Then GoTo ContinueLoop
        
        btn = value & number
        If emulatePiece(1) = btn Then
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
            playerTwo(emulatePiece(0))("danger") = True
            playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Exit For
        ElseIf buttons(btn)("isPiece") Then
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
        
        GoTo ContinueLoop
ContinueLoop:
    Next value

    ' Right
    For Each value In Array("A", "B", "C", "D", "E", "F", "G", "H")
        If indexLetter = 8 Then Exit For
        If numbers(value) > indexLetter Then
            btn = value & number
            If emulatePiece(1) = btn Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
                playerTwo(emulatePiece(0))("danger") = True
                playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
                Exit For
            ElseIf buttons(btn)("isPiece") Then
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
    
    If emulatePiece(0) <> "" And emulatePiece(1) <> "" Then
        playerTwo(emulatePiece(0))("newPos") = lastPos
        buttons(playerTwo(emulatePiece(0))("newPos"))("isPiece") = lastIsPiece
        buttons(playerTwo(emulatePiece(0))("newPos"))("player") = lastPlayer
        buttons(playerTwo(emulatePiece(0))("newPos"))("piece") = lastPiece
        buttons(emulatePiece(1))("isPiece") = lastIsPieceBtns
    End If

    If valuesAdded = 0 Then
        getPosRookP1 = Empty
    Else
        getPosRookP1 = availablePos.keys
    End If
End Function

Public Function getPosBishopP1(piece As String, Optional emulatePiece As Variant) As Variant
    getPosBishopP1 = Empty

    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Dim lastPos As Variant
    Dim lastPiece As Variant
    Dim lastPlayer As Integer
    Dim lastIsPiece As Boolean
    Dim lastIsPieceBtns As Boolean
    Set availablePos = CreateObject("Scripting.Dictionary")
    valuesAdded = 0

    If IsMissing(emulatePiece) Then emulatePiece = Array("", "")
    If emulatePiece(0) <> "" And emulatePiece(1) <> "" Then
        lastPos = playerTwo(emulatePiece(0))("newPos")
        lastPlayer = buttons(playerTwo(emulatePiece(0))("newPos"))("player")
        lastPiece = buttons(playerTwo(emulatePiece(0))("newPos"))("piece")
        lastIsPiece = buttons(playerTwo(emulatePiece(0))("newPos"))("isPiece")
        lastIsPieceBtns = buttons(emulatePiece(1))("isPiece")
        buttons(playerTwo(emulatePiece(0))("newPos"))("isPiece") = False
        buttons(playerTwo(emulatePiece(0))("newPos"))("player") = 0
        playerTwo(emulatePiece(0))("newPos") = emulatePiece(1)
        buttons(emulatePiece(1))("isPiece") = True
    End If

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
        If emulatePiece(1) = btn Then
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
            playerTwo(emulatePiece(0))("danger") = True
            playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Exit Do
        ElseIf buttons(btn)("isPiece") Then
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
        If emulatePiece(1) = btn Then
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
            playerTwo(emulatePiece(0))("danger") = True
            playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Exit Do
        ElseIf buttons(btn)("isPiece") Then
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
        If emulatePiece(1) = btn Then
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
            playerTwo(emulatePiece(0))("danger") = True
            playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Exit Do
        ElseIf buttons(btn)("isPiece") Then
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
        If emulatePiece(1) = btn Then
            availablePos.Add btn, True
            valuesAdded = valuesAdded + 1
            playerTwo(emulatePiece(0))("danger") = True
            playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Exit Do
        ElseIf buttons(btn)("isPiece") Then
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

    If emulatePiece(0) <> "" And emulatePiece(1) <> "" Then
        playerTwo(emulatePiece(0))("newPos") = lastPos
        buttons(playerTwo(emulatePiece(0))("newPos"))("isPiece") = lastIsPiece
        buttons(playerTwo(emulatePiece(0))("newPos"))("player") = lastPlayer
        buttons(playerTwo(emulatePiece(0))("newPos"))("piece") = lastPiece
        buttons(emulatePiece(1))("isPiece") = lastIsPieceBtns
    End If

    If valuesAdded = 0 Then Exit Function

    getPosBishopP1 = availablePos.keys
End Function
