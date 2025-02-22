Option Explicit

Dim frm As UserForm


Private Sub Init()
    Set frm = UserForms(0)
End Sub


Public Sub GenericClick(ByVal ctrl As MSForms.CommandButton)
    If activePiece = "" Then Exit Sub

    If frm Is Nothing Then Init

    If isPossibleMove(ctrl.name) Then
        movePiece ctrl.name, activePiece
        disablePiece activePiece
    End If
End Sub


Public Sub GenericClickLabel(ByVal ctrl As MSForms.label)
    If Not boolPlaying Then Exit Sub

    If frm Is Nothing Then Init
    Dim name As String, btn As String, number As String
    Dim maxValues As Integer
    name = ctrl.name
    number = Mid(name, 2, 1)
    '//! If name = "E1King" Then MsgBox CStr(isCheckMate("E1King", True))

    If activePiece <> "" Then
        If playerOneTurn And (number = "7" Or number = "8") Then
            btn = CStr(playerTwo(name)("newPos"))
            If isPossibleMove(btn) Then
                movePiece btn, activePiece
                disablePiece activePiece
            End If
        ElseIf Not playerOneTurn And (number = "1" Or number = "2") Then
            btn = CStr(playerOne(name)("newPos"))
            If isPossibleMove(btn) Then
                movePiece btn, activePiece
                disablePiece activePiece
            End If
        End If
    End If

    If playerOneTurn And (number = "1" Or number = "2") Then
        If activePiece <> "" Then disablePiece activePiece
        frm.Controls(name).BorderStyle = fmBorderStyleSingle
        activePiece = name
        paintCases True
    ElseIf Not playerOneTurn And (number = "7" Or number = "8") Then
        If activePiece <> "" Then disablePiece activePiece
        frm.Controls(name).BorderStyle = fmBorderStyleSingle
        activePiece = name
        paintCases False
    End If

End Sub