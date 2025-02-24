Option Explicit

Dim frm As UserForm
Dim pathAssets As String

Private Sub BtnBishop_Click()
    Dim piece As String
    piece = "Bishop"
    Init
    If playerOneTurn Then
        playerOne(activePiece)("type") = piece
        frm.Controls(activePiece).Picture = LoadPicture(pathAssets & piece & "Purple.jpg")
    Else
        playerTwo(activePiece)("type") = piece
        frm.Controls(activePiece).Picture = LoadPicture(pathAssets & piece & "White.jpg")
    End If
    Unload Me
End Sub

Private Sub BtnKnight_Click()
    Dim piece As String
    piece = "Knight"
    Init
    If playerOneTurn Then
        playerOne(activePiece)("type") = piece
        frm.Controls(activePiece).Picture = LoadPicture(pathAssets & piece & "Purple.jpg")
    Else
        playerTwo(activePiece)("type") = piece
        frm.Controls(activePiece).Picture = LoadPicture(pathAssets & piece & "White.jpg")
    End If
    Unload Me
End Sub

Private Sub BtnQueen_Click()
    Dim piece As String
    piece = "Queen"
    Init
    If playerOneTurn Then
        playerOne(activePiece)("type") = piece
        frm.Controls(activePiece).Picture = LoadPicture(pathAssets & piece & "Purple.jpg")
    Else
        playerTwo(activePiece)("type") = piece
        frm.Controls(activePiece).Picture = LoadPicture(pathAssets & piece & "White.jpg")
    End If
    Unload Me
End Sub

Private Sub BtnRook_Click()
    Dim piece As String
    piece = "Rook"
    Init
    If playerOneTurn Then
        playerOne(activePiece)("type") = piece
        frm.Controls(activePiece).Picture = LoadPicture(pathAssets & piece & "Purple.jpg")
    Else
        playerTwo(activePiece)("type") = piece
        frm.Controls(activePiece).Picture = LoadPicture(pathAssets & piece & "White.jpg")
    End If
    Unload Me
End Sub

Public Sub Init()
    If frm Is Nothing Then Set frm = UserForms(0)
End Sub

Private Sub UserForm_Initialize()
    pathAssets = pathGame & "assets\"
    If Not playerOneTurn Then
        BtnQueen.Picture = LoadPicture(pathAssets & "QueenPurple.jpg")
        BtnRook.Picture = LoadPicture(pathAssets & "RookPurple.jpg")
        BtnKnight.Picture = LoadPicture(pathAssets & "KnightPurple.jpg")
        BtnBishop.Picture = LoadPicture(pathAssets & "BishopPurple.jpg")
    Else
        BtnQueen.Picture = LoadPicture(pathAssets & "QueenWhite.jpg")
        BtnRook.Picture = LoadPicture(pathAssets & "RookWhite.jpg")
        BtnKnight.Picture = LoadPicture(pathAssets & "KnightWhite.jpg")
        BtnBishop.Picture = LoadPicture(pathAssets & "BishopWhite.jpg")
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "You cannot close this until you choose a piece.", vbExclamation, "Notice"
        Cancel = True
    End If
End Sub

