Option Explicit
Public boolPlaying As Boolean
Public playerOneTurn As Boolean
Public playerOne As Object
Public playerTwo As Object
Public buttons As Object
Public activePiece As String
Public pieceActived As String
Public letters As Object
Public numbers As Object
Public piecesEatenP1 As Integer
Public piecesEatenP2 As Integer
Public pathGame As String
Public colors As Object
Public lastMovement As Variant


Dim frm As UserForm
Dim i As Integer
Dim j As Integer
Dim value As Variant


Private Sub Init()
    Set frm = UserForms(0)
End Sub

Public Sub Delay(milliseconds As Single)
    Dim start As Single
    start = Timer
    Do While Timer < start + milliseconds / 1000
        DoEvents
    Loop
End Sub


Public Function isPossibleMove(button As String) As Boolean
    Dim i As Integer
    Dim value As Variant
    isPossibleMove = False
    Dim containsPos As Boolean
    Dim isPosibleEat As Boolean
    If playerOneTurn Then
        If IsEmpty(playerOne(activePiece)("nextPos")) Then Exit Function
        If activePiece = "E1King" Then
            If isCheck(activePiece, button, True) Then Exit Function
        End If
        If playerOne("E1King")("danger") Then
            If Not breaksCheck(activePiece, button, True) Then Exit Function
        End If
        containsPos = ArrayContains(playerOne(activePiece)("nextPos"), button)
        If Not containsPos Then Exit Function
        If activePiece <> "E1King" Then
            isPosibleEat = ArrayContains(playerOne(activePiece)("nextPos"), button)
            If isCheck("E1King", CStr(playerOne("E1King")("newPos")), True, Array(activePiece, button)) And Not isPosibleEat Then
                MsgBox "You cannot move there, you will be in check"
                Exit Function
            End If
        End If
    Else
        If IsEmpty(playerTwo(activePiece)("nextPos")) Then Exit Function
        If activePiece = "E8King" Then
            If isCheck(activePiece, button, False) Then Exit Function
        End If
        If playerTwo("E8King")("danger") Then
            If Not breaksCheck(activePiece, button, False) Then Exit Function
        End If
        containsPos = ArrayContains(playerTwo(activePiece)("nextPos"), button)
        If Not containsPos Then Exit Function
        If activePiece <> "E8King" Then
            isPosibleEat = ArrayContains(playerTwo(activePiece)("nextPos"), button)
            If isCheck("E8King", CStr(playerTwo("E8King")("newPos")), False, Array(activePiece, button)) And Not isPosibleEat Then
                MsgBox "You cannot move there, you will be in check"
                Exit Function
            End If
        End If
    End If
    isPossibleMove = True
End Function

Public Sub disablePiece(piece As String)
    If frm Is Nothing Then Init

    Dim value As Variant
    Dim value2 As Variant
    Dim newPosPiece As String
    Dim newPos As Object

    Set newPos = CreateObject("Scripting.Dictionary")
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
    pieceActived = activePiece
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
    kingP1 = "E1King"
    kingP2 = "E8King"
    rePaintCases

    If frm Is Nothing Then Init

    If Not IsEmpty(lastMovement) Then
        For Each value In lastMovement
            frm.Controls(value).BackColor = colors("lastMovement")
        Next value
    End If

    If boolPlayerOne Then
        availablePos = updateMoves(activePiece, True)
        If IsEmpty(availablePos) Then Exit Sub
        For Each pos In availablePos
            frm.Controls(pos).BackColor = colors("caseSelected")
        Next pos
    Else
        availablePos = updateMoves(activePiece, False)
        If IsEmpty(availablePos) Then Exit Sub
        For Each pos In availablePos
            frm.Controls(pos).BackColor = colors("caseSelected")
        Next pos
    End If
    
    If isCheck(kingP1, CStr(playerOne(kingP1)("newPos")), True) Then
        For Each value In playerOne(kingP1)("piecesEater")
            posPiece = playerTwo(value)("newPos")
            If frm.Controls(posPiece) <> buttons(posPiece)("bgcolor") Then
                frm.Controls(posPiece).BackColor = colors("pieceEaterAndCaseSelected")
            Else
                frm.Controls(posPiece).BackColor = colors("pieceEater")
            End If
        Next value
        frm.Controls(playerOne(kingP1)("newPos")).BackColor = colors("danger")
        If activePiece <> kingP1 Then Exit Sub
        For Each value In playerOne(kingP1)("nextPos")
            If isCheck(kingP1, CStr(value), True) Then
                frm.Controls(value).BackColor = colors("danger")
            End If
        Next value
    ElseIf isCheck(kingP2, CStr(playerTwo(kingP2)("newPos")), False) Then
        For Each value In playerTwo(kingP2)("piecesEater")
            posPiece = playerOne(value)("newPos")
            If frm.Controls(posPiece).BackColor <> colors("pieceEater") Then
                frm.Controls(posPiece).BackColor = colors("pieceEaterAndCaseSelected")
            Else
                frm.Controls(posPiece).BackColor = colors("pieceEater")
            End If
        Next value
        frm.Controls(playerTwo("E8King")("newPos")).BackColor = colors("danger")
        If activePiece <> kingP2 Then Exit Sub
        For Each value In playerTwo(kingP2)("nextPos")
            If isCheck(kingP2, CStr(value), False) Then
                frm.Controls(value).BackColor = colors("danger")
            End If
        Next value
    End If
    
    
End Sub

Public Function movePiece(button As String, piece As String)
    If frm Is Nothing Then Init
    Dim posBefore As String
    Dim pieceEaten As String
    
    If playerOneTurn Then
        If playerOne(piece)("type") = "Pawn" And Mid(button, 2, 1) = "8" Then PromotePawn.Show
        If piece = "E1King" Then
            If button = "G1" Then movePiece "F1", "H1Rook"
            If button = "C1" Then movePiece "D1", "A1Rook"
        End If
        lastMovement = Array(playerOne(activePiece)("newPos"), button)
        frm.Controls(piece).Left = CDbl(buttons(button)("posxy")("x")) + 5
        frm.Controls(piece).Top = CDbl(buttons(button)("posxy")("y")) + 5
        If buttons(button)("player") = 2 Then
            pieceEaten = buttons(button)("piece")
            playerTwo(pieceEaten)("dead") = True
            playerTwo(pieceEaten)("newPos") = ""
            If piecesEatenP1 > 7 Then
                frm.Controls(pieceEaten).Left = 390 + ((piecesEatenP1 - 8) * 20)
                frm.Controls(pieceEaten).Top = 287 + 40
            Else
                frm.Controls(pieceEaten).Left = 390 + (piecesEatenP1 * 20)
                frm.Controls(pieceEaten).Top = 287
            End If
            piecesEatenP1 = piecesEatenP1 + 1
        End If
        posBefore = playerOne(piece)("newPos")
        buttons(posBefore)("isPiece") = False
        buttons(posBefore)("player") = 0
        buttons(posBefore)("piece") = ""
        buttons(button)("isPiece") = True
        buttons(button)("player") = 1
        buttons(button)("piece") = piece
        playerOne(piece)("newPos") = button
        playerOne(piece)("moved") = True
        playerOne(piece)("firstMove") = False
    Else
        If playerTwo(piece)("type") = "Pawn" And Mid(button, 2, 1) = "1" Then PromotePawn.Show
        If piece = "E8King" Then
            If button = "G8" Then movePiece "F8", "H8Rook"
            If button = "C8" Then movePiece "D8", "A8Rook"
        End If
        lastMovement = Array(playerTwo(activePiece)("newPos"), button)
        frm.Controls(piece).Left = CDbl(buttons(button)("posxy")("x")) + 5
        frm.Controls(piece).Top = CDbl(buttons(button)("posxy")("y")) + 5
        If buttons(button)("player") = 1 Then
            pieceEaten = buttons(button)("piece")
            playerOne(pieceEaten)("dead") = True
            If piecesEatenP2 > 7 Then
                frm.Controls(pieceEaten).Left = 390 + ((piecesEatenP2 - 8) * 20)
                frm.Controls(pieceEaten).Top = 37
            Else
                frm.Controls(pieceEaten).Left = 390 + (piecesEatenP2 * 20)
                frm.Controls(pieceEaten).Top = 77
            End If
            piecesEatenP2 = piecesEatenP2 + 1
        End If
        posBefore = playerTwo(piece)("newPos")
        buttons(posBefore)("isPiece") = False
        buttons(posBefore)("player") = 0
        buttons(posBefore)("piece") = ""
        buttons(button)("isPiece") = True
        buttons(button)("player") = 2
        buttons(button)("piece") = piece
        playerTwo(piece)("newPos") = button
        playerTwo(activePiece)("moved") = True
        playerTwo(activePiece)("firstMove") = False
    End If
End Function



Public Function breaksCheck(piece As String, button As String, boolPlayerOne As Boolean) As Boolean
    breaksCheck = False
    If buttons(button)("piece") = "" Then
        If piece = "H7Pawn" Then MsgBox buttons(button)("piece") & buttons(button)("isPiece")
        If boolPlayerOne Then
            If activePiece = "E1King" And Not isCheck(activePiece, button, True) Then
                breaksCheck = True
                Exit Function
            End If
            If Not ArrayContains(playerOne(piece)("nextPos"), button) Then Exit Function
            If isCheck("E1King", CStr(playerOne("E1King")("newPos")), True, Array(piece, button)) Then
                MsgBox "You are in check right now, you cannot move that piece"
                Exit Function
            End If
            breaksCheck = True
        Else
            If activePiece = "E8King" And Not isCheck(activePiece, button, False) Then
                breaksCheck = True
                Exit Function
            End If
            If Not ArrayContains(playerTwo(piece)("nextPos"), button) Then Exit Function
            If isCheck("E8King", CStr(playerTwo("E8King")("newPos")), False, Array(piece, button)) Then
                MsgBox "You are in check right now, you cannot move that piece"
                Exit Function
            End If
            breaksCheck = True
        End If
        Exit Function

    End If
    

    Dim piecesEater As Variant
    Dim pieceToEat As Variant
    If boolPlayerOne Then
        piecesEater = deleteFromArr(playerOne(piece)("piecesEater"), buttons(button)("piece"))
        If Not IsEmpty(piecesEater) Then Exit Function
    Else
        piecesEater = deleteFromArr(playerTwo(piece)("piecesEater"), buttons(button)("piece"))
        If Not IsEmpty(piecesEater) Then Exit Function
    End If
    breaksCheck = True
End Function