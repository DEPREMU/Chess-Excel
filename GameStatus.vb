Option Explicit

Public boolCheckPlayer1 As Boolean
Public boolCheckPlayer2 As Boolean
Public gameFinished As Boolean

Dim frm As UserForm


Private Sub Init()
    Set frm = UserForms(0)
End Sub


Public Function checkGameStatus(piece As String)
    If frm Is Nothing Then Init
    Dim localPiece As Variant
    Dim kingP1 As String
    Dim kingP2 As String
    Dim isCheckP1 As Boolean
    Dim isCheckP2 As Boolean
    Dim isCheckMateP1 As Boolean
    Dim isCheckMateP2 As Boolean
    Dim piecesEater As Variant
    Dim currentPos As Variant
    kingP1 = "E1King"
    kingP2 = "E8King"
    
    If isStalemate() Then
        MsgBox "Stalemate, no possible valid movements"
        frm.LComments.Caption = "Stalemate"
        finishGame
        Exit Function
    End If

    If isInsufficientMaterial() Then
        MsgBox "Insufficient Material, the game finished"
        frm.LComments.Caption = "Insufficient Material"
        finishGame
        Exit Function
    End If
    
    isCheckMateP1 = isCheckMate(True)
    isCheckMateP2 = isCheckMate(False)
    
    If isCheckMateP1 Or isCheckMateP2 Or playerOne("E1King")("dead") Or _
            playerTwo("E8King")("dead") Then
        MsgBox "Player " & IIf(isCheckMateP1, "two", "one") & " wins"
        frm.LComments.Caption = "W Player " & IIf(isCheckMateP1, "Two", "One")
        currentPos = IIf(isCheckMateP1 Or playerOne("E1King")("dead"), playerOne(kingP1)("newPos"), _
            playerTwo(kingP2)("newPos"))
        If currentPos <> "" Then frm.Controls(currentPos).BackColor = colors("danger")
        isCheck CStr(currentPos), True
        piecesEater = IIf(isCheckMateP1, playerOne(kingP1)("piecesEater"), _
            playerTwo(kingP2)("piecesEater"))
        If Not IsEmpty(piecesEater) Then
            For Each localPiece In piecesEater
                currentPos = getPosPlayer(CStr(localPiece), Not isCheckMateP1)
                frm.Controls(currentPos).BackColor = colors("pieceEater")
            Next localPiece
        End If
        finishGame
        Exit Function
    End If

    isCheckP1 = isCheck(CStr(playerOne(kingP1)("newPos")), True)
    isCheckP2 = isCheck(CStr(playerTwo(kingP2)("newPos")), False)

    If Not isCheckP1 And Not isCheckP2 Then
        If Not playerOneTurn Then
            frm.LComments.Caption = "Player One Turn"
            boolCheckPlayer1 = False
        Else
            frm.LComments.Caption = "Player Two Turn"
            boolCheckPlayer2 = False
        End If
        rePaintCases
        Exit Function
    End If
    
    piecesEater = IIf(isCheckP1, playerOne(kingP1)("piecesEater"), _
        playerTwo(kingP2)("piecesEater"))
    
    frm.LComments.Caption = "Check Player " & IIf(isCheckP1, "One", "Two")
    For Each localPiece In piecesEater
        currentPos = getPosPlayer(CStr(localPiece), isCheckP2)
        frm.Controls(currentPos).BackColor = colors("pieceEater")
    Next localPiece
    currentPos = IIf(isCheckP1, playerOne(kingP1)("newPos"), _
        playerTwo(kingP2)("newPos"))
    frm.Controls(currentPos).BackColor = colors("danger")
    boolCheckPlayer1 = isCheckP1
    boolCheckPlayer2 = isCheckP2
End Function

Public Function isCheck(position As String, boolPlayerOne As Boolean, Optional emulatePiece As Variant)
    Dim pos As Variant
    Dim number As Integer
    Dim pieceP1 As Variant
    Dim pieceP2 As Variant
    Dim localPiece As Variant
    Dim indexLetter As Integer
    Dim availablePosP1 As Variant
    Dim availablePosP2 As Variant
    isCheck = False

    If IsMissing(emulatePiece) Then emulatePiece = Array("", "")

    If boolPlayerOne Then
        playerOne("E1King")("piecesEater") = Empty
        For Each pieceP2 In playerTwo.keys
            If Not playerTwo(pieceP2)("dead") Then
                If Mid(pieceP2, 2, 1) = "7" Then
                    pos = playerTwo(pieceP2)("newPos")
                    indexLetter = CInt(numbers(Mid(pos, 1, 1)))
                    number = CInt(Mid(pos, 2, 1))
                    availablePosP2 = getAvailablePosP2(CStr(pieceP2), emulatePiece)
                Else
                    availablePosP2 = getAvailablePosP2(CStr(pieceP2), emulatePiece)
                End If
                If Not IsEmpty(availablePosP2) Then
                    If ArrayContains(availablePosP2, position) Then
                        playerOne("E1King")("piecesEater") = addToArr(playerOne("E1King")("piecesEater"), pieceP2)
                        isCheck = True
                    End If
                End If
            End If
        Next pieceP2

    Else
        playerTwo("E8King")("piecesEater") = Empty
        For Each pieceP1 In playerOne.keys
            If Not playerOne(pieceP1)("dead") Then
                If Mid(pieceP1, 2, 1) = "2" Then
                    pos = playerOne(pieceP1)("newPos")
                    indexLetter = CInt(numbers(Mid(pos, 1, 1)))
                    number = CInt(Mid(pos, 2, 1))
                    availablePosP1 = getAvailablePosP1(CStr(pieceP1), emulatePiece)
                Else
                    availablePosP1 = getAvailablePosP1(CStr(pieceP1), emulatePiece)
                End If
                If Not IsEmpty(availablePosP1) Then
                    If ArrayContains(availablePosP1, position) Then
                        isCheck = True
                        playerTwo("E8King")("piecesEater") = addToArr(playerTwo("E8King")("piecesEater"), pieceP1)
                    End If
                End If
            End If
        Next pieceP1
    End If

End Function

Public Function isCheckMate(boolPlayerOne As Boolean) As Boolean
    Dim pos As Variant
    Dim value As Variant
    Dim eqValue As Variant
    Dim countPos As Integer
    Dim countMate As Integer
    Dim pieceEater As Variant
    Dim localPiece As Variant
    Dim availablePos As Variant
    Dim positionsEater As Variant
    Dim positionsPiece As Variant
    Dim isPossibleEat As Boolean
    Dim piecesPlayer As Variant
    Dim equalsValues As Variant
    Dim piecesEater As Variant
    Dim isDeadPiece As Boolean
    Dim currentPos As String
    Dim typePiece As String
    Dim boolPawn As Boolean
    Dim kingP1 As String
    Dim kingP2 As String
    kingP1 = "E1King"
    kingP2 = "E8King"
    countMate = 0
    countPos = 0
    isCheckMate = False


    If boolPlayerOne Then
        currentPos = CStr(playerOne(kingP1)("newPos"))
        availablePos = getAvailablePosP1(kingP1, currentPos)
        piecesEater = playerOne(kingP1)("piecesEater")
        piecesPlayer = deleteFromArr(playerOne.keys, "E1King")
    Else
        availablePos = getAvailablePosP2(kingP2, CStr(playerTwo(kingP2)("newPos")))
        currentPos = CStr(playerTwo(kingP2)("newPos"))
        piecesEater = playerTwo(kingP2)("piecesEater")
        piecesPlayer = deleteFromArr(playerTwo.keys, "E8King")
    End If

    If IsEmpty(availablePos) Or Not isCheck(currentPos, boolPlayerOne) Then Exit Function


    For Each pos In availablePos
        If isCheck(CStr(pos), boolPlayerOne) Then countMate = countMate + 1
        countPos = countPos + 1
    Next pos

    If isCheck(currentPos, boolPlayerOne) Then countMate = countMate + 1
    countPos = countPos + 1

    If Not IsEmpty(piecesEater) Then
        For Each pieceEater In piecesEater
            For Each localPiece In piecesPlayer
                isDeadPiece = getDeadState(CStr(localPiece), boolPlayerOne)

                If Not isDeadPiece Then
                    positionsEater = getNextPosPlayerNotByNextPos(CStr(pieceEater), Not boolPlayerOne)

                    positionsPiece = getNextPosPlayerNotByNextPos(CStr(localPiece), boolPlayerOne)


                    typePiece = getTypePiece(CStr(localPiece), boolPlayerOne)

                    If typePiece <> "Pawn" Then
                        If Not IsEmpty(positionsEater) And Not IsEmpty _
                                (positionsPiece) Then
                            equalsValues = equalsValuesArrs( _
                                addToArr(positionsEater, getPosPlayer(CStr(pieceEater), Not boolPlayerOne)), _
                                positionsPiece)
                            If Not IsEmpty(equalsValues) Then
                                For Each eqValue In equalsValues
                                    If Not isCheck(currentPos, boolPlayerOne, Array(localPiece, eqValue)) Then
                                        countMate = countMate - 1
                                    End If
                                Next eqValue
                            End If
                        End If
                    Else
                        If ArrayContains(positionsPiece, getPosPlayer(CStr(pieceEater), Not boolPlayerOne)) _
                                 Or breaksCheckPawn(CStr(localPiece), CStr(pieceEater), boolPlayerOne) Then
                            countMate = countMate - 1
                        End If
                    End If
                End If
            Next localPiece
        Next pieceEater
    End If

    If countMate >= countPos Then
        isCheckMate = True
        Exit Function
    End If
End Function

Public Function breaksCheckPawn(piece As String, pieceEater As String, boolCheckP1 As Boolean) As Boolean
    breaksCheckPawn = False

    Dim value As Variant
    Dim nextPos As Variant
    Dim currentPos As String
    Dim currentPosKing As String
    Dim nextPosEater As Variant
    Dim equalsValues As Variant
    Dim currentPosEater As String

    currentPos = getPosPlayer(piece, boolCheckP1)
    currentPosKing = getPosPlayer(IIf(boolCheckP1, "E1King", "E8King"), boolCheckP1)
    currentPosEater = getPosPlayer(pieceEater, Not boolCheckP1)

    If currentPos = currentPosEater Then
        breaksCheckPawn = True
        Exit Function
    End If
    
    nextPos = getNextPosPlayerNotByNextPos(piece, boolCheckP1)
    nextPosEater = getNextPosPlayerNotByNextPos(pieceEater, Not boolCheckP1)

    equalsValues = equalsValuesArrs(nextPosEater, nextPos)

    If IsEmpty(equalsValues) Then Exit Function
    
    For Each value In equalsValues
        If Not isCheck(currentPosKing, boolCheckP1, Array(piece, value)) Then
            breaksCheckPawn = True
            Exit Function
        End If
    Next value

End Function


Public Function isStalemate() As Boolean
    isStalemate = False
    Dim localPiece As Variant
    Dim newPos As Variant
    Dim pos As Variant

    If isCheck(CStr(playerOne("E1King")("newPos")), True) Or _
            isCheck(CStr(playerTwo("E8King")("newPos")), False) Then
        Exit Function
    End If

    For Each localPiece In playerOne.keys
        If Not playerOne(localPiece)("dead") Then
            newPos = getAvailablePosP1(CStr(localPiece))
            If Not IsEmpty(newPos) Then
                For Each pos In newPos
                    If Not isCheck(CStr(playerOne("E1King")("newPos")), True, Array(localPiece, pos)) Then Exit Function
                Next pos
            End If
        End If
    Next localPiece

    For Each localPiece In playerTwo.keys
        If Not playerTwo(localPiece)("dead") Then
            newPos = getAvailablePosP2(CStr(localPiece))
            If Not IsEmpty(newPos) Then
                For Each pos In newPos
                    If Not isCheck(CStr(playerTwo("E8King")("newPos")), False, Array(localPiece, pos)) Then Exit Function
                Next pos
            End If
        End If
    Next localPiece

    isStalemate = True
End Function

Public Function isInsufficientMaterial() As Boolean
    isInsufficientMaterial = False
    Dim piece As Variant
    Dim countPiecesP1 As Integer
    Dim countPiecesP2 As Integer
    countPiecesP1 = 0
    countPiecesP2 = 0

    For Each piece In playerOne.keys
        If Not playerOne(piece)("dead") Then countPiecesP1 = countPiecesP1 + 1
    Next piece

    For Each piece In playerTwo.keys
        If Not playerTwo(piece)("dead") Then countPiecesP2 = countPiecesP2 + 1
    Next piece
    
    If countPiecesP1 = 1 And countPiecesP2 = 1 Then
        isInsufficientMaterial = True
        Exit Function
    End If

    If (countPiecesP1 = 1 And countPiecesP2 = 2) Or (countPiecesP1 = 2 And countPiecesP2 = 1) Then
        ' King vs King and Knight
        ' King vs King and Bishop
        ' King and Bishop vs King
        ' King and Bishop vs King
        If playerTwo("B8Knight")("dead") And playerTwo("G8Knight")("dead") Then
            Exit Function
        ElseIf playerTwo("B8Bishop")("dead") And playerTwo("G8Bishop")("dead") Then
            Exit Function
        ElseIf playerOne("B1Knight")("dead") And playerOne("G1Knight")("dead") Then
            Exit Function
        ElseIf playerOne("B1Bishop")("dead") And playerOne("G1Bishop")("dead") Then
            Exit Function
        End If
    End If
    
    If countPiecesP1 = 2 And countPiecesP2 = 2 Then
        ' King and Bishop vs King and Bishop
        If (playerOne("C1Bishop")("dead") And playerTwo("F8Bishop")("dead")) Then
            isInsufficientMaterial = True
            Exit Function
        ElseIf (playerOne("F1Bishop")("dead") And playerTwo("C8Bishop")("dead")) Then
            isInsufficientMaterial = True
            Exit Function
        End If
    End If
End Function


Public Function finishGame()
    boolPlaying = False
    gameFinished = True
    changeStateButtons
End Function

