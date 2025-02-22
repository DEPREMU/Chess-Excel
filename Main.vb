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
    Set letters = CreateObject("Scripting.Dictionary")
    Set numbers = CreateObject("Scripting.Dictionary")
    i = 1
    For Each value In Array("A", "B", "C", "D", "E", "F", "G", "H")
        numbers.Add value, i
        letters.Add CStr(i), value
        i = i + 1
    Next value
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
    If playerOneTurn Then
        If IsEmpty(playerOne(activePiece)("nextPos")) Then
            isPossibleMove = False
            Exit Function
        End If
        If ArrayContains(playerOne(activePiece)("nextPos"), button) Then isPossibleMove = True
    Else
        If IsEmpty(playerTwo(activePiece)("nextPos")) Then
            isPossibleMove = False
            Exit Function
        End If
        If ArrayContains(playerTwo(activePiece)("nextPos"), button) Then isPossibleMove = True
    End If

    If Not isPossibleMove Then isPossibleMove = False
End Function

Public Sub disablePiece(piece As String)
    If frm Is Nothing Then Init

    Dim value As Variant
    Dim value2 As Variant
    Dim newPosPiece As String
    Dim newPos As Object

    Set newPos = CreateObject("Scripting.Dictionary")
    frm.Controls(piece).BorderStyle = fmBorderStyleNone


    If playerOneTurn Then
        If Not IsEmpty(playerOne(piece)("nextPos")) Then
            For Each value In playerOne(piece)("nextPos")
                frm.Controls(value).BackColor = buttons(value)("bgcolor")
            Next value
        End If
        If Not playerOne(piece)("moved") Then Exit Sub
        playerOne(piece)("nextPos") = getAvailablePosP1(piece)
        playerOne(piece)("moved") = False
        '//! swapLabels
    Else
        If Not IsEmpty(playerTwo(piece)("nextPos")) Then
            For Each value In playerTwo(piece)("nextPos")
                frm.Controls(value).BackColor = buttons(value)("bgcolor")
            Next value
        End If

        If Not playerTwo(piece)("moved") Then Exit Sub
        playerTwo(piece)("nextPos") = getAvailablePosP2(piece)
        playerTwo(piece)("moved") = False
        '//! swapLabels
    End If
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
    Dim value As Variant
    Dim values As Variant

    If frm Is Nothing Then Init

    If boolPlayerOne Then
        values = updateMoves(activePiece, True)
        If IsEmpty(values) Then Exit Sub
        For Each value In values
            If Mid(playerOne(activePiece)("newPos"), 1, 2) <> value Then frm.Controls(value).BackColor = &H80FFFF
        Next value
    Else
        values = updateMoves(activePiece, False)
        If IsEmpty(values) Then Exit Sub
        For Each value In values
            If Mid(playerTwo(activePiece)("newPos"), 1, 2) <> value Then frm.Controls(value).BackColor = &H80FFFF
        Next value
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
        checkGameStatus(piece)
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
        checkGameStatus(piece)
    End If
End Function