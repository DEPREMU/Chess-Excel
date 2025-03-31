Option Explicit


Public Function getNextPosBishop(piece As String, boolPlayerOne As Boolean, Optional emulatePiece As Variant) As Variant
    getNextPosBishop = Empty
    
    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Dim currentPos As String
    Dim i As Integer
    Dim j As Integer
    Dim lastPos As Variant
    Dim lastPiece As Variant
    Dim lastPlayer As Integer
    Dim lastIsPiece As Boolean
    Dim lastIsPieceBtns As Boolean
    Dim btn As String
    Dim player As Integer
    Dim localPiece As Variant
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
            If boolPlayerOne Then
                playerTwo(emulatePiece(0))("danger") = True
                playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Else
                playerOne(emulatePiece(0))("danger") = True
                playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
            End If
            Exit Do
        ElseIf buttons(btn)("isPiece") Then
            player = buttons(btn)("player")
            localPiece = buttons(btn)("piece")
            If player = 2 And boolPlayerOne Then
                playerTwo(localPiece)("danger") = True
                playerTwo(localPiece)("piecesEater") = addToArr(playerTwo(localPiece)("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            ElseIf player = 1 And Not boolPlayerOne Then
                playerOne(localPiece)("danger") = True
                playerOne(localPiece)("piecesEater") = addToArr(playerOne(localPiece)("piecesEater"), piece)
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
            If boolPlayerOne Then
                playerTwo(emulatePiece(0))("danger") = True
                playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Else
                playerOne(emulatePiece(0))("danger") = True
                playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
            End If
            Exit Do
        ElseIf buttons(btn)("isPiece") Then
            player = buttons(btn)("player")
            localPiece = buttons(btn)("piece")
            If player = 2 And boolPlayerOne Then
                playerTwo(localPiece)("danger") = True
                playerTwo(localPiece)("piecesEater") = addToArr(playerTwo(localPiece)("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            ElseIf player = 1 And Not boolPlayerOne Then
                playerOne(localPiece)("danger") = True
                playerOne(localPiece)("piecesEater") = addToArr(playerOne(localPiece)("piecesEater"), piece)
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
            If boolPlayerOne Then
                playerTwo(emulatePiece(0))("danger") = True
                playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Else
                playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
                playerOne(emulatePiece(0))("danger") = True
            End If
            Exit Do
        ElseIf buttons(btn)("isPiece") Then
            player = buttons(btn)("player")
            localPiece = buttons(btn)("piece")
            If player = 2 And boolPlayerOne Then
                playerTwo(localPiece)("danger") = True
                playerTwo(localPiece)("piecesEater") = addToArr(playerTwo(localPiece)("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            ElseIf player = 1 And Not boolPlayerOne Then
                playerOne(localPiece)("danger") = True
                playerOne(localPiece)("piecesEater") = addToArr(playerOne(localPiece)("piecesEater"), piece)
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
            If boolPlayerOne Then
                playerTwo(emulatePiece(0))("danger") = True
                playerTwo(emulatePiece(0))("piecesEater") = addToArr(playerTwo(emulatePiece(0))("piecesEater"), piece)
            Else
                playerOne(emulatePiece(0))("danger") = True
                playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
            End If
            Exit Do
        ElseIf buttons(btn)("isPiece") Then
            player = buttons(btn)("player")
            localPiece = buttons(btn)("piece")
            If player = 2 And boolPlayerOne Then
                playerTwo(localPiece)("danger") = True
                playerTwo(localPiece)("piecesEater") = addToArr(playerTwo(localPiece)("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            ElseIf player = 1 And Not boolPlayerOne Then
                playerOne(localPiece)("danger") = True
                playerOne(localPiece)("piecesEater") = addToArr(playerOne(localPiece)("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
            Exit Do
        End If
        availablePos.Add btn, True
        valuesAdded = valuesAdded + 1
    Loop
    
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
    
    If valuesAdded = 0 Then Exit Function
    
    getNextPosBishop = availablePos.keys
    Set availablePos = Nothing
End Function