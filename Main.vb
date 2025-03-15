Option Explicit

Public boolPlaying As Boolean
Public playerOneTurn As Boolean
Public playerOne As Object
Public playerTwo As Object
Public buttons As Object
Public activePiece As String
Public letters As Object
Public numbers As Object
Public piecesEatenP1 As Integer
Public piecesEatenP2 As Integer
Public pathGame As String
Public colors As Object
Public lastMovement As Variant
Public enPassant As Variant

Dim frm As UserForm
Dim i As Integer
Dim j As Integer
Dim value As Variant

Private Sub Init()
    Set frm = UserForms(0)
End Sub

Public Function isPossibleMove(button As String) As Boolean
    isPossibleMove = False
    
    Dim containsPos As Boolean
    Dim isPossibleEat As Boolean
    Dim nextPos As Variant
    Dim typePiece As String
    Dim kingInDanger As Boolean
    Dim kingPlayer As String
    Dim currentPos As String
    
    If playerOneTurn Then
        nextPos = playerOne(activePiece)("nextPos")
        typePiece = playerOne(activePiece)("type")
        kingInDanger = playerOne("E1King")("danger")
        kingPlayer = "E1King"
        currentPos = CStr(playerOne(activePiece)("newPos"))
    Else
        nextPos = playerTwo(activePiece)("nextPos")
        typePiece = playerTwo(activePiece)("type")
        kingInDanger = playerTwo("E8King")("danger")
        kingPlayer = "E8King"
        currentPos = CStr(playerTwo(activePiece)("newPos"))
    End If

    If IsEmpty(nextPos) Then Exit Function
    If typePiece = "King" Then
        If isCheck(button, playerOneTurn) Then Exit Function
    End If
    If kingInDanger Then
        If Not breaksCheck(activePiece, button, playerOneTurn) Then Exit Function
    End If
    containsPos = ArrayContains(nextPos, button)
    If Not containsPos Then Exit Function

    If typePiece <> "King" Then
        isPossibleEat = ArrayContains(nextPos, button)
        If isCheck(currentPos, playerOneTurn, Array(activePiece, button)) And Not isPossibleEat Then
            MsgBox "You cannot move there, you will be in check"
            Exit Function
        End If
    End If

    isPossibleMove = True
End Function

Public Sub disablePiece(piece As String)
    If frm Is Nothing Then Init

    frm.Controls(piece).BorderStyle = fmBorderStyleNone

    rePaintCases
    If playerOneTurn Then
        If Not playerOne(piece)("moved") Then Exit Sub
        playerOne(piece)("nextPos") = getAvailablePosP1(piece)
        playerOne(piece)("moved") = False
        '! swapLabels
    Else
        If Not playerTwo(piece)("moved") Then Exit Sub
        playerTwo(piece)("nextPos") = getAvailablePosP2(piece)
        playerTwo(piece)("moved") = False
        '! swapLabels
    End If
    checkGameStatus(piece)
    activePiece = ""
    playerOneTurn = Not playerOneTurn
End Sub

Public Function updateMoves(piece As String, boolPlayerOne As Boolean) As Variant
    Dim availablePos As Variant

    If boolPlayerOne Then
        availablePos = getAvailablePosP1(piece)
        playerOne(piece)("nextPos") = availablePos
    Else
        availablePos = getAvailablePosP2(piece)
        playerTwo(piece)("nextPos") = availablePos
    End If

    updateMoves = availablePos
End Function

Public Sub paintCases(boolPlayerOne As Boolean)
    Dim pos As Variant
    Dim availablePos As Variant
    Dim posPiece As Variant
    Dim kingP1 As String
    Dim kingP2 As String
    Dim typePiece As String
    Dim isCheckP1 As Boolean
    Dim isCheckP2 As Boolean
    Dim nextPos As Variant
    Dim piecesEater As Variant
    Dim currentPos As String
    kingP1 = "E1King"
    kingP2 = "E8King"
    rePaintCases

    If frm Is Nothing Then Init

    If Not IsEmpty(lastMovement) Then
        For Each value In lastMovement
            frm.Controls(value).BackColor = colors("lastMovement")
        Next value
    End If

    availablePos = updateMoves(activePiece, boolPlayerOne)
    If IsEmpty(availablePos) Then Exit Sub
    For Each pos In availablePos
        frm.Controls(pos).BackColor = colors("caseSelected")
    Next pos
    
    isCheckP1 = isCheck(CStr(playerOne(kingP1)("newPos")), True)
    isCheckP2 = isCheck(CStr(playerTwo(kingP2)("newPos")), False)
    
    If Not isCheckP1 And Not isCheckP2 Then Exit Sub
    
    If isCheckP1 Then
        piecesEater = playerOne(kingP1)("piecesEater")
        currentPos = CStr(playerOne(kingP1)("newPos"))
        typePiece = playerOne(activePiece)("type")
        nextPos = playerOne(kingP1)("nextPos")
    Else
        piecesEater = playerTwo(kingP2)("piecesEater")
        currentPos = CStr(playerTwo(kingP2)("newPos"))
        typePiece = playerTwo(activePiece)("type")
        nextPos = playerTwo(kingP2)("nextPos")
    End If
    
    For Each value In piecesEater
        posPiece = getPosPlayer(value, isCheckP2)
        If frm.Controls(posPiece) <> buttons(posPiece)("bgcolor") Then
            frm.Controls(posPiece).BackColor = colors("pieceEaterAndCaseSelected")
        Else
            frm.Controls(posPiece).BackColor = colors("pieceEater")
        End If
    Next value
    
    
    frm.Controls(currentPos).BackColor = colors("danger")
    If typePiece <> "King" Then Exit Sub
    
    For Each value In nextPos
        If isCheck(CStr(value), isCheckP1) Then
            frm.Controls(value).BackColor = colors("danger")
        End If
    Next value
    
End Sub


Public Function movePiece(button As String, piece As String)
    If frm Is Nothing Then Init
    Dim posBefore As String
    Dim pieceEaten As String
    Dim currentPos As String
    Dim typePiece As String
    Dim piecesEaten As Integer
    Dim isFirstMove As Boolean
    Dim boolMoreThan1 As Boolean
    Dim btn As String
    
    typePiece = getTypePiece(piece, playerOneTurn)
    If playerOneTurn Then
        isFirstMove = playerOne(piece)("firstMove")
        boolMoreThan1 = Mid(button, 2, 1) - Mid(playerOne(piece)("firstPos"), 2, 1) > 1
    Else
        isFirstMove = playerTwo(piece)("firstMove")
        boolMoreThan1 = Mid(button, 2, 1) + Mid(playerTwo(piece)("firstPos"), 2, 1) > 1
    End If
    
    If typePiece = "Pawn" Then
        If ArrayContains(Array("1", "8"), Mid(button, 2, 1)) Then
            PromotePawn.Show
            isCheck getPosPlayer(IIf( Not playerOneTurn, "E1King", "E8King"), Not playerOneTurn), Not playerOneTurn
        End If
        If Not IsEmpty(enPassant) Then
            If ArrayContains(Array(1, 2), enPassant(2)) And enPassant(1) = button Then
                pieceEaten = buttons(button)("enPassant")
                currentPos = getPosPlayer(piece, playerOneTurn)
                piecesEaten = IIf(playerOneTurn, piecesEatenP1, piecesEatenP2)
                buttons(button)("enPassant") = ""
                buttons(currentPos)("isPiece") = False
                buttons(currentPos)("player") = 0
                buttons(currentPos)("piece") = ""
                If playerOneTurn Then
                    playerTwo(pieceEaten)("dead") = True
                    playerTwo(pieceEaten)("newPos") = ""
                    piecesEatenP1 = piecesEatenP1 + 1
                Else
                    playerOne(pieceEaten)("dead") = True
                    playerOne(pieceEaten)("newPos") = ""
                    piecesEatenP2 = piecesEatenP2 + 1
                End If
                frm.Controls(pieceEaten).Left = 390 + ((piecesEaten - iif(piecesEaten > 7, 8, 0)) * 20)
                frm.Controls(pieceEaten).Top = iif(playerOneTurn, 287, 78) + iif(pieceEaten > 7, 40, 0)
                clearEnPassant
            End If
        End If
        
        If isFirstMove Then
            If boolMoreThan1 Then
                btn = Mid(piece, 1, 1) & CStr(CInt(Mid(piece, 2, 1)) + IIf(playerOneTurn, 1, - 1))
                MsgBox btn
                If Not buttons(btn)("isPiece") Then
                    buttons(btn)("enPassant") = activePiece
                    If playerOneTurn Then
                        playerOne(piece)("enPassant") = True
                    Else
                        playerTwo(piece)("enPassant") = True
                    End If
                    enPassant = Array(piece, btn, IIf(playerOneTurn, 1, 2))
                End If
            End If
        Else
            If playerOneTurn Then
                playerOne(piece)("enPassant") = False
            Else
                playerTwo(piece)("enPassant") = False
            End If
            clearEnPassant
        End If
    Else
        clearEnPassant
    End If

    If piece = "E1King" Then
        If button = "G1" Then movePiece "F1", "H1Rook"
        If button = "C1" Then movePiece "D1", "A1Rook"
    End If
    If piece = "E8King" Then
        If button = "G8" Then movePiece "F8", "H8Rook"
        If button = "C8" Then movePiece "D8", "A8Rook"
    End If
    lastMovement = Array(getPosPlayer(piece, playerOneTurn), button)
    frm.Controls(piece).Left = CDbl(buttons(button)("posxy")("x")) + 5
    frm.Controls(piece).Top = CDbl(buttons(button)("posxy")("y")) + 5
    addPieceToEatenPieces(button)
    clearButtons piece, button


End Function

Public Function breaksCheck(piece As String, button As String, boolPlayerOne As Boolean) As Boolean
    Dim typePiece As String
    Dim nextPos As Variant
    Dim currentPos As String
    Dim piecesEater As Variant

    breaksCheck = False

    If boolPlayerOne Then
        typePiece = playerOne(piece)("type")
        nextPos = playerOne(piece)("nextPos")
        currentPos = CStr(playerOne("E1King")("newPos"))
        piecesEater = deleteFromArr(playerOne(piece)("piecesEater"), buttons(button)("piece"))
    Else
        typePiece = playerTwo(piece)("type")
        nextPos = playerTwo(piece)("nextPos")
        currentPos = CStr(playerTwo("E8King")("newPos"))
        piecesEater = deleteFromArr(playerTwo(piece)("piecesEater"), buttons(button)("piece"))
    End If

    If buttons(button)("piece") = "" Then
        If typePiece = "King" Then
            If Not isCheck(button, boolPlayerOne) Then
                breaksCheck = True
                Exit Function
            End If
        End If
        If Not ArrayContains(nextPos, button) Then Exit Function
        If isCheck(currentPos, boolPlayerOne, Array(piece, button)) Then
            MsgBox "You are in check right now, you cannot move that piece"
            Exit Function
        End If
        breaksCheck = True
        Exit Function
    End If

    If Not IsEmpty(piecesEater) Then Exit Function

    breaksCheck = True
End Function


Public Function clearEnPassant()
    If Not IsEmpty(enPassant) Then
        buttons(enPassant(1))("enPassant") = ""
        If enPassant(2) = 1 Then
            playerOne(enPassant(0))("enPassant") = False
        Else
            playerTwo(enPassant(0))("enPassant") = False
        End If
        enPassant = Empty
    End If
End Function


Public Function clearButtons(piece As String, button As Variant)
    Dim posBefore As String
    posBefore = getPosPlayer(piece, playerOneTurn)
    buttons(posBefore)("isPiece") = False
    buttons(posBefore)("player") = 0
    buttons(posBefore)("piece") = ""
    buttons(button)("isPiece") = True
    buttons(button)("player") = IIf(playerOneTurn, 1, 2)
    buttons(button)("piece") = piece
    If playerOneTurn Then
        playerOne(piece)("newPos") = button
        playerOne(piece)("moved") = True
        playerOne(piece)("firstMove") = False
    Else
        playerTwo(piece)("newPos") = button
        playerTwo(piece)("moved") = True
        playerTwo(piece)("firstMove") = False
    End If
End Function

Public Function addPieceToEatenPieces(button As String)
    Dim piecesEaten As Integer
    Dim pieceEaten As Variant
    Dim pieceEater As Variant
    pieceEater = IIf(playerOneTurn, piecesEatenP1, piecesEatenP2)
    pieceEaten = buttons(button)("piece")
    If pieceEaten = "" Then Exit Function

    If buttons(button)("player") = 2 Then
        playerTwo(pieceEaten)("dead") = True
        playerTwo(pieceEaten)("newPos") = ""
        piecesEatenP1 = piecesEatenP1 + 1
    ElseIf buttons(button)("player") = 1 Then
        playerOne(pieceEaten)("dead") = True
        playerOne(pieceEaten)("newPos") = ""
        piecesEatenP2 = piecesEatenP2 + 1
    End If

    frm.Controls(pieceEaten).Left = 390 + ((piecesEaten - iif(piecesEaten > 7, 8, 0)) * 20)
    frm.Controls(pieceEaten).Top = iif(playerOneTurn, 287, 78) + iif(piecesEaten > 7, 40, 0)
End Function