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
    
    
    If isCheckMate("E1King", True) Then
        MsgBox "Player two wins"
        frm.Comments.Caption = "W Player Two"
        frm.Controls(playerOne("E1King")("newPos")).BackColor = &H80 &
        finishGame
        Exit Function
    ElseIf isCheckMate("E8King", False) Then
        MsgBox "Player one wins"
        frm.Comments.Caption = "W Player One"
        frm.Controls(playerTwo("E8King")("newPos")).BackColor = &H80 &
        finishGame
        Exit Function
    End If
    
    If isCheck("E1King", CStr(playerOne("E1King")("newPos")), True) Then
        frm.Comments.Caption = "Check Player One"
        For Each localPiece In playerOne("E1King")("piecesEater")
            frm.Controls(playerTwo(localPiece)("newPos")).BackColor = &HFFFF80
        Next localPiece
        frm.Controls(playerOne("E1King")("newPos")).BackColor = &H80 &
        boolCheckPlayer1 = True
    ElseIf isCheck("E8King", CStr(playerTwo("E8King")("newPos")), False) Then
        For Each localPiece In playerTwo("E8King")("piecesEater")
            frm.Controls(playerOne(localPiece)("newPos")).BackColor = &HFFFF80
        Next localPiece
        frm.Comments.Caption = "Check Player Two"
        frm.Controls(playerTwo("E8King")("newPos")).BackColor = &H80 &
        boolCheckPlayer2 = True
    Else
        If Not playerOneTurn Then
            frm.Comments.Caption = "Player One Turn"
            boolCheckPlayer1 = False
            rePaintCases
        Else
            frm.Comments.Caption = "Player Two Turn"
            boolCheckPlayer2 = False
            rePaintCases
        End If
    End If
    
    
End Function


Public Function isCheck(piece As String, position As String, boolPlayerOne As Boolean, Optional debugg As Boolean)
    Dim pos As Variant
    Dim number As String
    Dim pieceP1 As Variant
    Dim pieceP2 As Variant
    Dim localPiece As Variant
    Dim indexLetter As String
    Dim availablePosP1 As Variant
    Dim availablePosP2 As Variant
    isCheck = False

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
                availablePosP2 = getAvailablePosP2(CStr(pieceP2))
            End If
            If IsEmpty(availablePosP2) Then Goto ContinueLoopP1
            If ArrayContains(availablePosP2, position) Then
                playerOne("E1King")("piecesEater") = addToArr(playerOne("E1King")("piecesEater"), pieceP2)
                isCheck = True
            End If

            ContinueLoopP1 :
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
                availablePosP1 = getAvailablePosP1(CStr(pieceP1))
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
    Dim countPos As Integer
    Dim countMate As Integer
    Dim pos As Variant
    countMate = 0
    countPos = 0
    
    If boolPlayerOne Then
        availablePos = possiblePosKingP1(piece, CStr(playerOne(piece)("newPos")))
        If IsEmpty(availablePos) Then
            isCheckMate = False
            Exit Function
        End If
        If isCheck(piece, CStr(playerOne(piece)("newPos")), True) Then countMate = countMate + 1
        countPos = countPos + 1
        For Each pos In availablePos
            If isCheck(piece, CStr(pos), True) Then countMate = countMate + 1
            countPos = countPos + 1
        Next pos
        If countPos = countMate Then
            For Each pos In availablePos
                frm.Controls(pos).BackColor = &H80 &
            Next pos
            isCheckMate = True
            Exit Function
        End If
    Else
        availablePos = posiblePosKingP2(piece, CStr(playerTwo(piece)("newPos")))
        If IsEmpty(availablePos) Then
            isCheckMate = False
            Exit Function
        End If
        If isCheck(piece, CStr(playerTwo(piece)("newPos")), False) Then countMate = countMate + 1
        countPos = countPos + 1
        For Each pos In availablePos
            If isCheck(piece, CStr(pos), False) Then countMate = countMate + 1
            countPos = countPos + 1
        Next pos
        If countPos = countMate Then
            isCheckMate = True
            Exit Function
        End If
    End If
    isCheckMate = False
End Function


Public Function finishGame()
    If boolPlaying Then boolPlaying = False
    boolCheckPlayer1 = False
    boolCheckPlayer2 = False
    gameFinished = True
    handleButtons
End Function