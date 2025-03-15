Option Explicit

Dim frm As UserForm
Dim pathAssets As String

Private Sub BtnBishop_Click()
    Dim piece As String
    piece = "Bishop"
    changePiece (piece)
    Unload Me
End Sub

Private Sub BtnKnight_Click()
    Dim piece As String
    piece = "Knight"
    changePiece (piece)
    Unload Me
End Sub

Private Sub BtnQueen_Click()
    Dim piece As String
    piece = "Queen"
    changePiece (piece)
    Unload Me
End Sub

Private Sub BtnRook_Click()
    Dim piece As String
    piece = "Rook"
    changePiece (piece)
    Unload Me
End Sub

Private Function changePiece(piece As String)
    If playerOneTurn Then
        playerOne(activePiece)("type") = piece
        frm.Controls(activePiece).Picture = LoadPicture(pathAssets & piece & "Purple.jpg")
    Else
        playerTwo(activePiece)("type") = piece
        frm.Controls(activePiece).Picture = LoadPicture(pathAssets & piece & "White.jpg")
    End If
End Function

Public Sub Init()
    If frm Is Nothing Then Set frm = UserForms(0)
End Sub

Private Sub UserForm_Initialize()
    Dim pieces As Variant
    Dim color As Variant
    Dim piece As Variant
    Init
    pieces = Array("Queen", "Rook", "Knight", "Bishop")
    color = IIf(Not playerOneTurn, "White", "Purple")
    pathAssets = pathGame & "assets\"

    For Each piece In pieces
        Me.Controls("Btn" & piece).Picture = LoadPicture(pathAssets & _
            piece & color & ".jpg")
    Next piece
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "You cannot close this until you choose a piece.", vbExclamation, "Notice"
        Cancel = True
    End If
End Sub

