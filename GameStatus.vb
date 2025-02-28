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
    kingP1 = "E1King"
    kingP2 = "E8King"

    If isStalemate() Then
        MsgBox "Stalemate, no posible valid movements"
        frm.LComments.Caption = "Stalemate"
        finishGame
        Exit Function
    End If
    
    If isInnsuficientMaterial() Then
        MsgBox "Innsuficient Material, the game finished"
        frm.LComments.Caption = "Innsuficient Material"
        finishGame
        Exit Function
    End If

    If isCheckMate(kingP1, True) Then
        MsgBox "Player two wins"
        frm.LComments.Caption = "W Player Two"
        frm.Controls(playerOne(kingP1)("newPos")).BackColor = colors("danger")
        isCheck kingP1, CStr(playerOne(kingP1)("newPos")), True
        For Each localPiece In playerOne(kingP1)("piecesEater")
            frm.Controls(playerTwo(localPiece)("newPos")).BackColor = colors("pieceEater")
        Next localPiece
        finishGame
        Exit Function
    ElseIf isCheckMate(kingP2, False) Then
        MsgBox "Player one wins"
        frm.LComments.Caption = "W Player One"
        isCheck kingP2, CStr(playerTwo(kingP2)("newPos")), False
        For Each localPiece In playerTwo(kingP2)("piecesEater")
            frm.Controls(playerOne(localPiece)("newPos")).BackColor = colors("pieceEater")
        Next localPiece
        frm.Controls(playerTwo(kingP2)("newPos")).BackColor = colors("danger")
        finishGame
        Exit Function
    End If

    If isCheck("E1King", CStr(playerOne("E1King")("newPos")), True) Then
        frm.LComments.Caption = "Check Player One"
        For Each localPiece In playerOne("E1King")("piecesEater")
            frm.Controls(playerTwo(localPiece)("newPos")).BackColor = colors("pieceEater")
        Next localPiece
        frm.Controls(playerOne("E1King")("newPos")).BackColor = colors("danger")
        boolCheckPlayer1 = True
    ElseIf isCheck("E8King", CStr(playerTwo("E8King")("newPos")), False) Then
        For Each localPiece In playerTwo("E8King")("piecesEater")
            frm.Controls(playerOne(localPiece)("newPos")).BackColor = colors("pieceEater")
        Next localPiece
        frm.LComments.Caption = "Check Player Two"
        frm.Controls(playerTwo("E8King")("newPos")).BackColor = colors("danger")
        boolCheckPlayer2 = True
    Else
        If playerOneTurn Then
            frm.LComments.Caption = "Player One Turn"
            boolCheckPlayer1 = False
            rePaintCases
        Else
            frm.LComments.Caption = "Player Two Turn"
            boolCheckPlayer2 = False
            rePaintCases
        End If
    End If


End Function


Public Function isCheck(piece As String, position As String, boolPlayerOne As Boolean, Optional emulatePiece As Variant, Optional debugg As Boolean)
    Dim pos As Variant
    Dim number As String
    Dim pieceP1 As Variant
    Dim pieceP2 As Variant
    Dim localPiece As Variant
    Dim indexLetter As String
    Dim availablePosP1 As Variant
    Dim availablePosP2 As Variant
    isCheck = False
    
    If IsMissing(emulatePiece) Then emulatePiece = Array("", "")
    If IsMissing(debugg) Then debugg = False
    
    
    
    If boolPlayerOne Then
        playerOne("E1King")("piecesEater") = Empty
        For Each pieceP2 In playerTwo.keys
            If playerTwo(pieceP2)("dead") Then Goto ContinueLoopP1
            If Mid(pieceP2, 2, 1) = "7" Then
                pos = playerTwo(pieceP2)("newPos")
                indexLetter = numbers(Mid(pos, 1, 1))
                number = Mid(pos, 2, 1)
                availablePosP2 = Array(letters(CStr(CInt(indexLetter) - 1)) & CStr(CInt(number) - 1), letters(CStr(CInt(indexLetter) + 1)) & CStr(CInt(number) - 1))
            Else
                availablePosP2 = getAvailablePosP2(CStr(pieceP2), emulatePiece)
            End If
            If IsEmpty(availablePosP2) Then Goto ContinueLoopP1
            If ArrayContains(availablePosP2, position) Then
                playerOne("E1King")("piecesEater") = addToArr(playerOne("E1King")("piecesEater"), pieceP2)
                isCheck = True
            End If
            
            ContinueLoopP1 :
            If debugg Then MsgBox CStr(pieceP2) & " |   " & Join(availablePosP2, ", ") & "  |  " & Join(emulatePiece, ", ")
            
        Next pieceP2
        
    Else
        playerTwo("E8King")("piecesEater") = Empty
        For Each pieceP1 In playerOne.keys
            If playerOne(pieceP1)("dead") Then Goto ContinueLoopP2
            If Mid(pieceP1, 2, 1) = "2" Then
                pos = playerOne(pieceP1)("newPos")
                indexLetter = numbers(Mid(pos, 1, 1))
                number = Mid(pos, 2, 1)
                availablePosP1 = Array(letters(CStr(CInt(indexLetter) - 1)) & CStr(CInt(number) + 1), letters(CStr(CInt(indexLetter) + 1)) & CStr(CInt(number) + 1))
            Else
                availablePosP1 = getAvailablePosP1(CStr(pieceP1), emulatePiece)
            End If
            If IsEmpty(availablePosP1) Then Goto ContinueLoopP2
            If ArrayContains(availablePosP1, position) Then
                isCheck = True
                playerTwo("E8King")("piecesEater") = addToArr(playerTwo("E8King")("piecesEater"), pieceP1)
            End If
            
            ContinueLoopP2 :
        Next pieceP1
    End If
    
End Function


Public Function isCheckMate(piece As String, boolPlayerOne As Boolean) As Boolean
    Dim availablePos As Variant
    Dim availablePos2 As Variant
    Dim countPos As Integer
    Dim countMate As Integer
    Dim pos As Variant
    Dim pieceEater As Variant
    Dim value As Variant
    Dim positionsEater As Variant
    Dim positionsPiece As Variant
    Dim localPiece As Variant
    Dim eqValue As Variant
    Dim equalsValues As Variant
    countMate = 0
    countPos = 0
    isCheckMate = False

    If boolPlayerOne Then
        availablePos = getPosKingP1(piece, CStr(playerOne(piece)("newPos")))
        If IsEmpty(availablePos) Or Not isCheck("E1King", CStr(playerOne("E1King")("newPos")), True) Then Exit Function
        For Each pos In availablePos
            If isCheck(piece, CStr(pos), True) Then countMate = countMate + 1
            countPos = countPos + 1
        Next pos
        If isCheck(piece, CStr(playerOne(piece)("newPos")), True) Then countMate = countMate + 1
        countPos = countPos + 1
        If Not IsEmpty(playerOne("E1King")("piecesEater")) Then
            For Each pieceEater In playerOne("E1King")("piecesEater")
                For Each localPiece In playerOne.keys
                    positionsEater = getAvailablePosP2(CStr(pieceEater))
                    positionsPiece = getAvailablePosP1(CStr(localPiece))
                    If Not IsEmpty(positionsEater) And Not IsEmpty(positionsPiece) And Not playerOne(localPiece)("dead") Then
                        equalsValues = equalsValuesArrs(addToArr(positionsEater, playerTwo(pieceEater)("newPos")), positionsPiece)
                        If Not IsEmpty(equalsValues) Then
                            For Each eqValue In equalsValues
                                If Not isCheck("E1King", CStr(playerOne("E1King")("newPos")), True, Array(localPiece, eqValue)) Then
                                    countMate = countMate - 1
                                ElseIf eqValue = playerTwo(pieceEater)("newPos") Then
                                    countMate = countMate - 1
                                End If
                            Next eqValue
                        End If
                    End If
                Next localPiece
            Next pieceEater
        End If
        If countMate >= countPos Then
            isCheckMate = True
            Exit Function
        End If
        
    Else
        availablePos = posiblePosKingP2(piece, CStr(playerTwo(piece)("newPos")))
        If IsEmpty(availablePos) Or Not isCheck("E8King", CStr(playerTwo("E8King")("newPos")), False) Then Exit Function
        For Each pos In availablePos
            If isCheck(piece, CStr(pos), False) Then countMate = countMate + 1
            countPos = countPos + 1
        Next pos
        If isCheck(piece, CStr(playerTwo(piece)("newPos")), False) Then countMate = countMate + 1
        countPos = countPos + 1
        If Not IsEmpty(playerTwo("E8King")("piecesEater")) Then
            For Each pieceEater In playerTwo("E8King")("piecesEater")
                For Each localPiece In playerTwo.keys
                    positionsEater = getAvailablePosP1(CStr(pieceEater))
                    positionsPiece = getAvailablePosP2(CStr(localPiece))
                    If Not IsEmpty(positionsEater) And Not IsEmpty(positionsPiece) And Not playerTwo(localPiece)("dead") Then
                        equalsValues = equalsValuesArrs(addToArr(positionsEater, playerOne(pieceEater)("newPos")), positionsPiece)
                        If Not IsEmpty(equalsValues) Then
                            For Each eqValue In equalsValues
                                If Not isCheck("E8King", CStr(playerTwo("E8King")("newPos")), False, Array(localPiece, eqValue)) Then
                                    countMate = countMate - 1
                                ElseIf eqValue = playerTwo(localPiece)("newPos") Then
                                    countMate = countMate - 1
                                End If
                            Next eqValue
                        End If
                    End If
                Next localPiece
            Next pieceEater
        End If
        If countMate >= countPos Then
            isCheckMate = True
            Exit Function
        End If
    End If
End Function

Public Function isStalemate() As Boolean
    isStalemate = False
    Dim localPiece As Variant
    Dim newPos As Variant
    Dim pos As Variant
    
    If isCheck("E1King", CStr(playerOne("E1King")("newPos")), True) Or _
            isCheck("E8King", CStr(playerTwo("E8King")("newPos")), False) Then
        Exit Function
    End If
    
    For Each localPiece In playerOne.keys
        If Not playerOne(localPiece)("dead") Then
            newPos = getAvailablePosP1(CStr(localPiece))
            If Not IsEmpty(newPos) Then
                For Each pos In newPos
                    If Not isCheck("E1King", CStr(playerOne("E1King")("newPos")), True, Array(localPiece, pos)) Then Exit Function
                Next pos
            End If
        End If
    Next localPiece
    
    For Each localPiece In playerTwo.keys
        If Not playerTwo(localPiece)("dead") Then
            newPos = getAvailablePosP2(CStr(localPiece))
            If Not IsEmpty(newPos) Then
                For Each pos In newPos
                    If Not isCheck("E8King", CStr(playerTwo("E8King")("newPos")), False, Array(localPiece, pos)) Then Exit Function
                Next pos
            End If
        End If
    Next localPiece
    
    isStalemate = True
End Function

Public Function isInnsuficientMaterial()
    isInnsuficientMaterial = False
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
        isInnsuficientMaterial = True
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
        If Not (playerOne("C1Bishop")("dead") And playerTwo("F8Bishop")("dead")) Then
            Exit Function
        ElseIf Not (playerOne("F1Bishop")("dead") And playerTwo("C8Bishop")("dead")) Then
            Exit Function
        End If
    End If
    isInnsuficientMaterial = True
End Function


Public Function finishGame()

    boolPlaying = False
    gameFinished = True
    handleButtons
End Function