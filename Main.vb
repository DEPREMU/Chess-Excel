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
    If playerOneTurn Then
        If IsEmpty(playerOne(activePiece)("nextPos")) Then Exit Function
        If activePiece = "E1King" Then
            If isCheck(activePiece, button, True) Then Exit Function
        End If
        If playerOne("E1King")("danger") Then
            If Not breaksCheck(activePiece, button, True) Then Exit Function
            isPossibleMove = True
            Exit Function
        End If
        
        If ArrayContains(playerOne(activePiece)("nextPos"), button) Then isPossibleMove = True
    Else
        If IsEmpty(playerTwo(activePiece)("nextPos")) Then Exit Function
        If activePiece = "E8King" Then
            If isCheck(activePiece, button, False) Then Exit Function
        End If
        If playerTwo("E8King")("danger") Then
            If Not breaksCheck(activePiece, button, False) Then Exit Function
            isPossibleMove = True
            Exit Function
        End If

        If ArrayContains(playerTwo(activePiece)("nextPos"), button) Then isPossibleMove = True
    End If
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
    pieceActived = activePiece
    activePiece = ""
    playerOneTurn = Not playerOneTurn
    checkGameStatus(piece)


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

    If frm Is Nothing Then Init

    If boolPlayerOne Then
        availablePos = updateMoves(activePiece, True)
        If IsEmpty(availablePos) Then Exit Sub
        For Each pos In availablePos
            If activePiece = "E1King" And isCheck(activePiece, CStr(pos), True) Then
                frm.Controls(pos).BackColor = &H80 &
            Else
                frm.Controls(pos).BackColor = &H80FFFF
            End If
        Next pos
    Else
        availablePos = updateMoves(activePiece, False)
        If IsEmpty(availablePos) Then Exit Sub
        For Each pos In availablePos
            If activePiece = "E8King" And isCheck(activePiece, CStr(pos), False) Then
                frm.Controls(pos).BackColor = &H80 &
            Else
                frm.Controls(pos).BackColor = &H80FFFF
            End If
        Next pos
    End If
    
End Sub

Public Function movePiece(button As String, piece As String)
    If frm Is Nothing Then Init
    Dim posBefore As String
    Dim pieceEaten As String
    
    If playerOneTurn Then
        If piece = "E1King" Then
            If button = "G1" Then movePiece "F1", "H1Rook"
            If button = "C1" Then movePiece "D1", "A1Rook"
        End If
        frm.Controls(piece).Left = CDbl(buttons(button)("posxy")("x")) + 5
        frm.Controls(piece).Top = CDbl(buttons(button)("posxy")("y")) + 5
        If buttons(button)("player") = 2 Then
            pieceEaten = buttons(button)("piece")
            playerTwo(pieceEaten)("dead") = True
            If piecesEatenP1 > 5 Then
                frm.Controls(pieceEaten).Left = 400 + ((piecesEatenP1 - 6) * 25)
                frm.Controls(pieceEaten).Top = 287 + 40
            Else
                frm.Controls(pieceEaten).Left = 400 + (piecesEatenP1 * 25)
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
        If piece = "E8King" Then
            If button = "G8" Then movePiece "F8", "H8Rook"
            If button = "C8" Then movePiece "D8", "A8Rook"
        End If
        frm.Controls(piece).Left = CDbl(buttons(button)("posxy")("x")) + 5
        frm.Controls(piece).Top = CDbl(buttons(button)("posxy")("y")) + 5
        If buttons(button)("player") = 1 Then
            pieceEaten = buttons(button)("piece")
            playerOne(pieceEaten)("dead") = True
            If piecesEatenP2 > 5 Then
                frm.Controls(pieceEaten).Left = 400 + ((piecesEatenP2 - 6) * 25)
                frm.Controls(pieceEaten).Top = 77 + 40
            Else
                frm.Controls(pieceEaten).Left = 400 + (piecesEatenP2 * 25)
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
    If Not buttons(button)("isPiece") Then
        If boolPlayerOne Then
            If activePiece = "E1King" And Not isCheck(activePiece, button, True) Then
                breaksCheck = True
                Exit Function
            End If
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