Option Explicit

Dim frm As UserForm

Private Sub Init()
    Set frm = UserForms(0)
End Sub

Public Sub posPieces()
    Dim value As Variant
    If frm Is Nothing Then Init

    For Each value In playerOne.keys
        If playerOne(value)("dead") Then Goto ContinueLoop
        frm.Controls(value).Left = frm.Controls(playerOne(value)("newPos")).Left + 5
        frm.Controls(value).Top = frm.Controls(playerOne(value)("newPos")).Top + 5

        ContinueLoop :
    Next value

    For Each value In playerTwo.keys
        If playerTwo(value)("dead") Then Goto ContinueLoop1
        frm.Controls(value).Left = frm.Controls(playerTwo(value)("newPos")).Left + 5
        frm.Controls(value).Top = frm.Controls(playerTwo(value)("newPos")).Top + 5

        ContinueLoop1 :
    Next value
End Sub

Public Function rePosPieces()
    If frm Is Nothing Then Init
    Dim piece As Variant
    
    For Each piece In playerOne.keys
        frm.Controls(piece).Left = CInt(buttons(Mid(piece, 1, 2))("posxy")("x")) + 5
        frm.Controls(piece).Top = CInt(buttons(Mid(piece, 1, 2))("posxy")("y")) + 5
    Next piece
    For Each piece In playerTwo.keys
        frm.Controls(piece).Left = buttons(Mid(piece, 1, 2))("posxy")("x") + 5
        frm.Controls(piece).Top = buttons(Mid(piece, 1, 2))("posxy")("y") + 5
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
    Dim letter As String
    Dim number As String
    Dim value As Variant
    Dim value2 As Variant

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
    posPieces

End Sub


Public Sub handleButtons()
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
            If playerOne("E1King")("newPos") = value And boolCheckPlayer1 Then Goto ContinueLoop
            If playerTwo("E8King")("newPos") = value And boolCheckPlayer2 Then Goto ContinueLoop
            frm.Controls(value).BackColor = buttons(value)("bgcolor")
        End If
        ContinueLoop :
    Next value
    
End Function