Option Explicit
Dim buttonHandlers() As New clsButtonHandler

Dim letter As Variant, possibleNextPositionP1 As Object
Dim possibleNextPositionP2 As Object
Dim chessPiece1 As Object, chessPiece2 As Object, buttonPosition As Object
Dim possibleNextEater As Object, value As Variant
Dim boolIsPiece As Boolean, pieceP1_1 As Object, pieceP1_2 As Object
Dim pieceP2_1 As Object, pieceP2_2 As Object


Private Sub ButtonStartGame_Click()
    If boolPlaying Or gameFinished Then Exit Sub
    handleButtons
    boolPlaying = True
    LComments.Caption = "Player one turn"
    boolCheckPlayer1 = False
    boolCheckPlayer2 = False
    gameFinished = False
End Sub

Private Sub ButtonRestartGame_Click()
    If Not boolPlaying And Not gameFinished Then Exit Sub
    '! If Not playerOneTurn Then swapLabels
    handleButtons
    boolPlaying = False
    If activePiece <> "" Then disablePiece activePiece
    activePiece = ""
    rePosPieces
    initializeGame
    LComments.Caption = "Game restarted"
    boolCheckPlayer1 = False
    boolCheckPlayer1 = False
    rePaintCases
End Sub

Private Sub UserForm_Initialize()
    initializeGame
End Sub


Public Function initializeGame()
    boolPlaying = False
    gameFinished = False
    playerOneTurn = True
    pathGame = "C:\Users\zae47\OneDrive\Documentos\ArchivosParaLaUniBIS\Tareas\Tetra4\AppsDesign\Ajedrez\Chess-Excel\"
    piecesEatenP1 = 0
    piecesEatenP2 = 0
    lastMovement = Empty

    Set playerOne = CreateObject("Scripting.Dictionary")
    Set playerTwo = CreateObject("Scripting.Dictionary")
    Set buttons = CreateObject("Scripting.Dictionary")
    Set letters = CreateObject("Scripting.Dictionary")
    Set numbers = CreateObject("Scripting.Dictionary")
    Set colors = CreateObject("Scripting.Dictionary")

    Dim isPiece As Object, posButton As Object, buttonCompleted As Object
    Dim piece1 As String, piece2 As String, buttonLocal As String
    Dim typePiece As String, possibleNextPositionP1_1 As Object
    Dim possibleNextPositionP1_2 As Object, possibleNextPositionP2_1 As Object
    Dim possibleNextPositionP2_2 As Object, ctrl As control, piece3 As String
    Dim piece4 As String
    Dim letter As Variant
    Dim i As Integer

    colors.Add "danger", &H33FF
    colors.Add "caseSelected", &HFFD700
    colors.Add "pieceEaterAndCaseSelected", &H80FF&
    colors.Add "pieceEater", &HFF6347
    colors.Add "BlackCase", RGB(125, 135, 150)
    colors.Add "WhiteCase", RGB(240, 217, 181)
    colors.Add "lastMovement", &HFFC0FF
    
    ' x = IIF(xx) Mod 2 = 0 RGB(255,200,150), RGB(194, 117, 25)


    i = 1
    For Each value In Array("A", "B", "C", "D", "E", "F", "G", "H")
        numbers.Add value, i
        letters.Add CStr(i), value
        i = i + 1
    Next value
    

    ' Make buttons info -------------------------------------------------------
    For Each letter In Array("A", "B", "C", "D", "E", "F", "G", "H")
        For i = 1 To 8
            buttonLocal = letter & CStr(i)
            Set buttonCompleted = CreateObject("Scripting.Dictionary")
            Set posButton = CreateObject("Scripting.Dictionary")
            posButton.Add "x", Controls(buttonLocal).Left
            posButton.Add "y", Controls(buttonLocal).Top

            buttonCompleted.Add "isPiece", False
            buttonCompleted.Add "posxy", posButton
            buttonCompleted.Add "bgcolor", colors("WhiteCase")
            buttonCompleted.Add "player", 0
            buttonCompleted.Add "name", buttonLocal

            Dim bool1 As Boolean, bool2 As Boolean
            bool1 = (letter = "A" Or letter = "C" Or letter = "E" Or letter = "G") And (i = 1 Or i = 3 Or i = 5 Or i = 7)
            bool2 = (letter = "B" Or letter = "D" Or letter = "F" Or letter = "H") And (i = 2 Or i = 4 Or i = 6 Or i = 8)
            If bool1 Or bool2 Then
                buttonCompleted("bgcolor") = colors("BlackCase")
            End If
            buttons.Add buttonLocal, buttonCompleted
            Controls(buttonLocal).Enabled = False
            Controls(buttonLocal).ZOrder(1)
        Next i
    Next letter

    ' Pawns pieces -------------------------------------------------------------
    For Each letter In Array("A", "B", "C", "D", "E", "F", "G", "H")
        Set chessPiece1 = CreateObject("Scripting.Dictionary")
        Set chessPiece2 = CreateObject("Scripting.Dictionary")
        Set possibleNextPositionP1 = CreateObject("Scripting.Dictionary")
        Set possibleNextPositionP2 = CreateObject("Scripting.Dictionary")

        possibleNextPositionP1.Add "1", letter & "3"
        possibleNextPositionP1.Add "2", letter & "4"
        possibleNextPositionP2.Add "1", letter & "6"
        possibleNextPositionP2.Add "2", letter & "5"

        piece1 = CStr(letter) & "2"
        piece2 = CStr(letter) & "7"

        buttons(piece1)("isPiece") = True
        buttons(piece2)("isPiece") = True
        buttons(piece1)("player") = 1
        buttons(piece2)("player") = 2
        buttons(piece1)("piece") = piece1 & "Pawn"
        buttons(piece2)("piece") = piece2 & "Pawn"

        chessPiece1.Add "firstPos", piece1
        chessPiece2.Add "firstPos", piece2
        chessPiece1.Add "newPos", piece1
        chessPiece2.Add "newPos", piece2
        chessPiece1.Add "nextPos", possibleNextPositionP1.items
        chessPiece2.Add "nextPos", possibleNextPositionP2.items
        chessPiece1.Add "type", "Pawn"
        chessPiece2.Add "type", "Pawn"
        chessPiece1.Add "moved", False
        chessPiece2.Add "moved", False
        chessPiece1.Add "firstMove", True
        chessPiece2.Add "firstMove", True
        chessPiece1.Add "danger", False
        chessPiece2.Add "danger", False
        chessPiece1.Add "piecesEater", Empty
        chessPiece2.Add "piecesEater", Empty
        chessPiece1.Add "dead", False
        chessPiece2.Add "dead", False

        playerOne.Add piece1 & "Pawn", chessPiece1
        playerTwo.Add piece2 & "Pawn", chessPiece2
    Next letter

    ' Pieces A->C And F->H -----------------------------------------------------
    For i = 1 To 3
        Set possibleNextPositionP1_1 = CreateObject("Scripting.Dictionary")
        Set possibleNextPositionP1_2 = CreateObject("Scripting.Dictionary")
        Set possibleNextPositionP2_1 = CreateObject("Scripting.Dictionary")
        Set possibleNextPositionP2_2 = CreateObject("Scripting.Dictionary")
        Set pieceP1_1 = CreateObject("Scripting.Dictionary")
        Set pieceP1_2 = CreateObject("Scripting.Dictionary")
        Set pieceP2_1 = CreateObject("Scripting.Dictionary")
        Set pieceP2_2 = CreateObject("Scripting.Dictionary")

        If i = 1 Then
            typePiece = "Rook"
            ' Player one
            piece1 = "A1"
            piece2 = "H1"
            ' Player two
            piece3 = "A8"
            piece4 = "H8"
        ElseIf i = 2 Then
            typePiece = "Knight"
            ' Player one
            piece1 = "B1"
            piece2 = "G1"
            possibleNextPositionP1_1.Add "1", "A3"
            possibleNextPositionP1_1.Add "2", "C3"
            possibleNextPositionP1_2.Add "1", "F3"
            possibleNextPositionP1_2.Add "2", "H3"

            ' Player two
            piece3 = "B8"
            piece4 = "G8"
            possibleNextPositionP2_1.Add "1", "A6"
            possibleNextPositionP2_1.Add "2", "C6"
            possibleNextPositionP2_2.Add "1", "F6"
            possibleNextPositionP2_2.Add "2", "H6"
        ElseIf i = 3 Then
            typePiece = "Bishop"
            ' Player one
            piece1 = "C1"
            piece2 = "F1"
            'Player two
            piece3 = "C8"
            piece4 = "F8"
        End If

        ' Player one
        pieceP1_1.Add "firstPos", piece1
        pieceP1_2.Add "firstPos", piece2
        pieceP1_1.Add "newPos", piece1
        pieceP1_2.Add "newPos", piece2
        pieceP1_1.Add "nextPos", possibleNextPositionP1_1.items
        pieceP1_2.Add "nextPos", possibleNextPositionP1_2.items
        pieceP1_1.Add "type", typePiece
        pieceP1_2.Add "type", typePiece
        pieceP1_1.Add "moved", False
        pieceP1_2.Add "moved", False
        pieceP1_1.Add "firstMove", True
        pieceP1_2.Add "firstMove", True
        pieceP1_1.Add "piecesEater", Empty
        pieceP1_2.Add "piecesEater", Empty
        pieceP1_1.Add "dead", False
        pieceP1_2.Add "dead", False

        ' Player two
        pieceP2_1.Add "firstPos", piece3
        pieceP2_2.Add "firstPos", piece4
        pieceP2_1.Add "newPos", piece3
        pieceP2_2.Add "newPos", piece4
        pieceP2_1.Add "nextPos", possibleNextPositionP2_1.items
        pieceP2_2.Add "nextPos", possibleNextPositionP2_2.items
        pieceP2_1.Add "type", typePiece
        pieceP2_2.Add "type", typePiece
        pieceP2_1.Add "moved", False
        pieceP2_2.Add "moved", False
        pieceP2_1.Add "firstMove", True
        pieceP2_2.Add "firstMove", True
        pieceP2_1.Add "piecesEater", Empty
        pieceP2_2.Add "piecesEater", Empty
        pieceP2_1.Add "dead", False
        pieceP2_2.Add "dead", False


        buttons(piece1)("isPiece") = True
        buttons(piece2)("isPiece") = True
        buttons(piece3)("isPiece") = True
        buttons(piece4)("isPiece") = True

        buttons(piece1)("piece") = piece1 & typePiece
        buttons(piece2)("piece") = piece2 & typePiece
        buttons(piece3)("piece") = piece3 & typePiece
        buttons(piece4)("piece") = piece4 & typePiece

        buttons(piece1)("player") = 1
        buttons(piece2)("player") = 1
        buttons(piece3)("player") = 2
        buttons(piece4)("player") = 2

        playerOne.Add piece1 & typePiece, pieceP1_1
        playerOne.Add piece2 & typePiece, pieceP1_2

        playerTwo.Add piece3 & typePiece, pieceP2_1
        playerTwo.Add piece4 & typePiece, pieceP2_2

    Next i

    ' Queens -------------------------------------------------------------------

    Set chessPiece1 = CreateObject("Scripting.Dictionary")
    Set chessPiece2 = CreateObject("Scripting.Dictionary")
    Set possibleNextPositionP1 = CreateObject("Scripting.Dictionary")
    Set possibleNextPositionP2 = CreateObject("Scripting.Dictionary")

    piece1 = "D1"
    piece2 = "D8"

    buttons(piece1)("isPiece") = True
    buttons(piece2)("isPiece") = True
    buttons(piece1)("player") = 1
    buttons(piece2)("player") = 2
    buttons(piece1)("piece") = piece1 & "Queen"
    buttons(piece2)("piece") = piece2 & "Queen"

    chessPiece1.Add "firstPos", piece1
    chessPiece2.Add "firstPos", piece2
    chessPiece1.Add "newPos", piece1
    chessPiece2.Add "newPos", piece2
    chessPiece1.Add "nextPos", possibleNextPositionP1.items
    chessPiece2.Add "nextPos", possibleNextPositionP2.items
    chessPiece1.Add "type", "Queen"
    chessPiece2.Add "type", "Queen"
    chessPiece1.Add "moved", False
    chessPiece2.Add "moved", False
    chessPiece1.Add "firstMove", True
    chessPiece2.Add "firstMove", True
    chessPiece1.Add "piecesEater", Empty
    chessPiece2.Add "piecesEater", Empty


    playerOne.Add piece1 & "Queen", chessPiece1
    playerTwo.Add piece2 & "Queen", chessPiece2

    ' King ---------------------------------------------------------------------
    Set chessPiece1 = CreateObject("Scripting.Dictionary")
    Set chessPiece2 = CreateObject("Scripting.Dictionary")
    Set possibleNextPositionP1 = CreateObject("Scripting.Dictionary")
    Set possibleNextPositionP2 = CreateObject("Scripting.Dictionary")

    piece1 = "E1"
    piece2 = "E8"

    buttons(piece1)("isPiece") = True
    buttons(piece2)("isPiece") = True
    buttons(piece1)("player") = 1
    buttons(piece2)("player") = 2
    buttons(piece1)("piece") = piece1 & "King"
    buttons(piece2)("piece") = piece2 & "King"

    chessPiece1.Add "firstPos", piece1
    chessPiece2.Add "firstPos", piece2
    chessPiece1.Add "newPos", piece1
    chessPiece2.Add "newPos", piece2
    chessPiece1.Add "nextPos", possibleNextPositionP1.items
    chessPiece2.Add "nextPos", possibleNextPositionP2.items
    chessPiece1.Add "type", "King"
    chessPiece2.Add "type", "King"
    chessPiece1.Add "moved", False
    chessPiece2.Add "moved", False
    chessPiece1.Add "firstMove", True
    chessPiece2.Add "firstMove", True
    chessPiece1.Add "danger", False
    chessPiece2.Add "danger", False
    chessPiece1.Add "piecesEater", Empty
    chessPiece2.Add "piecesEater", Empty
    chessPiece1.Add "dead", False
    chessPiece2.Add "dead", False

    playerOne.Add piece1 & "King", chessPiece1
    playerTwo.Add piece2 & "King", chessPiece2


    ' Make controls dinamically -----------------------------------------------
    i = 0
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CommandButton" Then
            ReDim Preserve buttonHandlers(i)
            Set buttonHandlers(i) = New clsButtonHandler
            Set buttonHandlers(i).btn = ctrl
            i = i + 1
        ElseIf TypeName(ctrl) = "Label" Then
            ReDim Preserve buttonHandlers(i)
            Set buttonHandlers(i) = New clsButtonHandler
            Set buttonHandlers(i).label = ctrl
            i = i + 1
        End If
    Next ctrl


    rePaintCases

End Function