Option Explicit

Dim frm As UserForm

Private Sub Init()
    Set frm = UserForms(0)
End Sub

Public Sub placePieces()
    Dim value As Variant
    Dim currentPos As Variant
    If frm Is Nothing Then Init

    For Each value In playerOne.keys
        If Not playerOne(value)("dead") Then
            currentPos = playerOne(value)("newPos")
            frm.Controls(value).Left = frm.Controls(currentPos).Left + 5
            frm.Controls(value).Top = frm.Controls(currentPos).Top + 5
        End If
    Next value

    For Each value In playerTwo.keys
        If Not playerTwo(value)("dead") Then
            currentPos = playerTwo(value)("newPos")
            frm.Controls(value).Left = frm.Controls(currentPos).Left + 5
            frm.Controls(value).Top = frm.Controls(currentPos).Top + 5
        End If
    Next value
End Sub

Public Function repositionPieces()
    If frm Is Nothing Then Init
    Dim piece As Variant
    Dim buttonPiece As Variant
    
    For Each piece In playerOne.keys
        buttonPiece = Mid(piece, 1, 2)
        frm.Controls(piece).Left = CInt(buttons(buttonPiece)("posxy")("x")) + 5
        frm.Controls(piece).Top = CInt(buttons(buttonPiece)("posxy")("y")) + 5
    Next piece
    For Each piece In playerTwo.keys
        buttonPiece = Mid(piece, 1, 2)
        frm.Controls(piece).Left = buttons(buttonPiece)("posxy")("x") + 5
        frm.Controls(piece).Top = buttons(buttonPiece)("posxy")("y") + 5
    Next piece
End Function

Public Sub swapLabels()
    If frm Is Nothing Then Init

    Dim i As Integer
    Dim letters As String
    Dim numbers As String

    If frm.LA.Caption = "A" Then
        letters = "HGFEDCBA"
        numbers = "87654321"
    Else
        letters = "ABCDEFGH"
        numbers = "12345678"
    End If

    For i = 1 To 8
        frm.Controls("L" & Chr(64 + i)).Caption = Mid(letters, i, 1)
        frm.Controls("L" & i).Caption = Mid(numbers, i, 1)
    Next i
    swapButtons
End Sub

Public Sub swapButtons()
    If frm Is Nothing Then Init

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim value As Variant

    If frm.LA.Caption = "H" Then
        Dim z As Integer
        z = 8
        For Each value In letters.items
            i = 1
            j = 8
            For k = 1 To 8
                frm.Controls(value & CStr(i)).Left = buttons(letters(CStr(z)) & CStr(j))("posxy")("x")
                frm.Controls(value & CStr(i)).Top = buttons(letters(CStr(z)) & CStr(j))("posxy")("y")
                i = i + 1
                j = j - 1
            Next k
            z = z - 1
        Next value
    Else
        For Each value In buttons.keys
            frm.Controls(value).Left = buttons(value)("posxy")("x")
            frm.Controls(value).Top = buttons(value)("posxy")("y")
        Next value
    End If
    placePieces
End Sub

Public Sub changeStateButtons()
    If frm Is Nothing Then Init
    Dim button As Variant

    For Each button In buttons.keys
        frm.Controls(button).Enabled = Not frm.Controls(button).Enabled
    Next button
End Sub

Public Function rePaintCases()
    If frm Is Nothing Then Init
    
    Dim value As Variant
    For Each value In buttons.keys
        If frm.Controls(value).BackColor <> buttons(value)("bgcolor") Then
            If playerOne("E1King")("newPos") = value _
                 And boolCheckPlayer1 Then Goto ContinueLoop
            If playerTwo("E8King")("newPos") = value _
                 And boolCheckPlayer2 Then Goto ContinueLoop
            frm.Controls(value).BackColor = buttons(value)("bgcolor")
        End If
        ContinueLoop :
    Next value
    
End Function

Public Function getPosPlayer(value As Variant, boolPlayerOne As Boolean) As Variant
    If boolPlayerOne Then
        getPosPlayer = playerOne(value)("newPos")
    Else
        getPosPlayer = playerTwo(value)("newPos")
    End If
End Function

Public Function getNextPosPlayer(value As Variant, boolPlayerOne As Boolean) As Variant
    If boolPlayerOne Then
        getNextPosPlayer = playerOne(value)("nextPos")
    Else
        getNextPosPlayer = playerTwo(value)("nextPos")
    End If
End Function