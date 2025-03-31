Option Explicit


Public Function getNextPosPawn(piece As String, boolPlayerOne As Boolean) As Variant
    Dim playerPiece As String
    Dim enemyEnPassant As Object
    Dim firstMove As Boolean
    Dim playerNumber As Integer
    Dim enemyNumber As Integer
    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim btn As String
    Dim pos As String
    Dim availablePos As Object
    Set availablePos = CreateObject("Scripting.Dictionary")
    Dim leftOrRight As Integer
    Dim i As Integer
    
    ' Usar `boolPlayerOne` para asignar el jugador correspondiente
    If boolPlayerOne Then
        ' Jugador 1
        playerPiece = playerOne(piece)("newPos")
        Set enemyEnPassant = playerTwo
        firstMove = playerOne(piece)("firstMove")
        playerNumber = 1
        enemyNumber = 2
    Else
        ' Jugador 2
        playerPiece = playerTwo(piece)("newPos")
        Set enemyEnPassant = playerOne
        firstMove = playerTwo(piece)("firstMove")
        playerNumber = 2
        enemyNumber = 1
    End If
    
    letter = Mid(playerPiece, 1, 1)
    indexLetter = numbers(letter)
    number = Mid(playerPiece, 2, 1)
    
    ' Loop to check for left and right diagonals (en passant)
    For i = 0 To 1
        leftOrRight = IIf(i = 0, 1, -1)
        If letters(CStr(indexLetter - leftOrRight)) <> "" Then
            btn = letters(CStr(indexLetter - leftOrRight)) & number
            If btn <> "" Then
                If buttons(btn)("player") = enemyNumber Then
                    If enemyEnPassant(buttons(btn)("piece"))("enPassant") Then
                        pos = enemyEnPassant(buttons(btn)("piece"))("newPos")
                        availablePos.Add Mid(pos, 1, 1) & CStr(CInt(Mid(pos, 2, 1)) + IIf(boolPlayerOne, 1, -1)), True
                    End If
                End If
            End If
        End If
    Next i

    ' Move forward logic
    If CInt(number) < 8 And boolPlayerOne Or CInt(number) > 1 And Not boolPlayerOne Then
        btn = letter & CStr(CInt(number) + IIf(boolPlayerOne, 1, -1))
        If Not buttons(btn)("isPiece") Then
            availablePos.Add btn, True
            If firstMove Then
                If Not buttons(letter & CStr(CInt(number) + IIf(boolPlayerOne, 2, -2)))("isPiece") Then
                    availablePos.Add letter & CStr(CInt(number) + IIf(boolPlayerOne, 2, -2)), True
                End If
            End If
        End If
    End If
    
    ' Capture diagonally left
    If indexLetter > 1 And CInt(number) < 8 And boolPlayerOne Or indexLetter > 1 And CInt(number) > 1 And Not boolPlayerOne Then
        btn = letters(CStr(indexLetter - 1)) & CStr(CInt(number) + IIf(boolPlayerOne, 1, -1))
        If buttons(btn)("player") = enemyNumber Then
            availablePos.Add btn, True
            enemyEnPassant(buttons(btn)("piece"))("danger") = True
            enemyEnPassant(buttons(btn)("piece"))("piecesEater") = addToArr(enemyEnPassant(buttons(btn)("piece"))("piecesEater"), piece)
        End If
    End If
    
    ' Capture diagonally right
    If indexLetter < 8 And CInt(number) < 8 And boolPlayerOne Or indexLetter < 8 And CInt(number) > 1 And Not boolPlayerOne Then
        btn = letters(CStr(indexLetter + 1)) & CStr(CInt(number) + IIf(boolPlayerOne, 1, -1))
        If buttons(btn)("player") = enemyNumber Then
            enemyEnPassant(buttons(btn)("piece"))("danger") = True
            enemyEnPassant(buttons(btn)("piece"))("piecesEater") = addToArr(enemyEnPassant(buttons(btn)("piece"))("piecesEater"), piece)
            availablePos.Add btn, True
        End If
    End If
    
    If availablePos.count = 0 Then
        getNextPosPawn = Empty
    Else
        getNextPosPawn = availablePos.keys
    End If
    Set availablePos = Nothing
End Function
