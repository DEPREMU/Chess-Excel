Option Explicit
Dim buttonHandlers() As New clsButtonHandler

Dim letter As Variant
Dim value As Variant

Private Sub ButtonStartGame_Click()
    If boolPlaying Or gameFinished Then Exit Sub
    changeStateButtons
    boolPlaying = True
    LComments.Caption = "Player one turn"
    boolCheckPlayer1 = False
    boolCheckPlayer2 = False
    gameFinished = False
End Sub

Private Sub ButtonRestartGame_Click()
    If Not boolPlaying And Not gameFinished Then Exit Sub
    '! If Not playerOneTurn Then swapLabels
    changeStateButtons
    boolPlaying = False
    If activePiece <> "" Then disablePiece activePiece
    activePiece = ""
    repositionPieces
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
    pathGame = _
        "C:\Users\zae47\OneDrive\Documentos\ArchivosParaLaUniBIS\Tareas\Tetra4\AppsDesign\Ajedrez\Chess-Excel\"
    piecesEatenP1 = 0
    piecesEatenP2 = 0
    lastMovement = Empty
    
    Set playerOne = CreateObject("Scripting.Dictionary")
    Set playerTwo = CreateObject("Scripting.Dictionary")
    Set buttons = CreateObject("Scripting.Dictionary")
    Set letters = CreateObject("Scripting.Dictionary")
    Set numbers = CreateObject("Scripting.Dictionary")
    Set colors = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    Dim ctrl As control
    Dim piece as variant
    Dim piece1 As String
    Dim piece2 As String
    Dim piece3 As String
    Dim piece4 As String
    Dim letter As Variant
    Dim isPiece As Object
    Dim pieceP1_1 As Object
    Dim pieceP1_2 As Object
    Dim pieceP2_1 As Object
    Dim pieceP2_2 As Object
    Dim posButton As Object
    Dim buttonLocal As String
    Dim chessPiece1 As Object
    Dim chessPiece2 As Object
    Dim buttonCompleted As Object
    Dim possibleNextPositionP1 As variant
    Dim possibleNextPositionP2 As variant
    Dim possibleNextPositionP1_1 As variant
    Dim possibleNextPositionP1_2 As variant
    Dim possibleNextPositionP2_1 As variant
    Dim possibleNextPositionP2_2 As variant
    
    
    colors.Add "danger", &H33FF
    colors.Add "caseSelected", &HFFD700
    colors.Add "pieceEaterAndCaseSelected", &H80FF &
    colors.Add "pieceEater", &HFF6347
    colors.Add "BlackCase", RGB(125, 135, 150)
    colors.Add "WhiteCase", RGB(240, 217, 181)
    colors.Add "lastMovement", &HFFC0FF

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
            buttonCompleted.Add "enPassant", ""
            
            Dim bool1 As Boolean
            Dim bool2 As Boolean
            bool1 = ArrayContains(Array("A", "C", "E", "G"), letter) _
                 And ArrayContains(Array(1, 3, 5, 7), i)
            bool2 = ArrayContains(Array("B", "D", "F", "H"), letter) _
                 And ArrayContains(Array(2, 4, 6, 8), i)
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
        possibleNextPositionP1 = Array(letter & "3", letter & "4")
        possibleNextPositionP2 = Array(letter & "6", letter & "5")
        
        piece1 = CStr(letter) & "2"
        piece2 = CStr(letter) & "7"
        
        For Each value In Array(piece1, piece2)
            buttons(value)("isPiece") = True
            buttons(value)("player") = iif(value = piece1, 1, 2)
            buttons(value)("piece") = value & "Pawn"
        Next value

        chessPiece1.Add "firstPos", piece1
        chessPiece2.Add "firstPos", piece2
        chessPiece1.Add "newPos", piece1
        chessPiece2.Add "newPos", piece2
        chessPiece1.Add "nextPos", possibleNextPositionP1
        chessPiece2.Add "nextPos", possibleNextPositionP2
        chessPiece1.Add "type", "Pawn"
        chessPiece2.Add "type", "Pawn"
        chessPiece1.Add "piecesEater", Empty
        chessPiece2.Add "piecesEater", Empty

        For Each value In Array("firstMove", "moved", "danger", "dead", "enPassant")
            chessPiece1.Add value, iif(value = "firstMove", True, False)
            chessPiece2.Add value, iif(value = "firstMove", True, False)
        Next value
        
        playerOne.Add piece1 & "Pawn", chessPiece1
        playerTwo.Add piece2 & "Pawn", chessPiece2
    Next letter
    
    ' Pieces A->C And F->H -----------------------------------------------------
    For Each piece In Array("Rook", "Knight", "Bishop")
        possibleNextPositionP1_1 = Empty
        possibleNextPositionP1_2 = Empty
        possibleNextPositionP2_1 = Empty
        possibleNextPositionP2_2 = Empty
        Set pieceP1_1 = CreateObject("Scripting.Dictionary")
        Set pieceP1_2 = CreateObject("Scripting.Dictionary")
        Set pieceP2_1 = CreateObject("Scripting.Dictionary")
        Set pieceP2_2 = CreateObject("Scripting.Dictionary")
        
        If piece = "Rook" Then
            ' Player one
            piece1 = "A1"
            piece2 = "H1"
            ' Player two
            piece3 = "A8"
            piece4 = "H8"
        ElseIf piece = "Knight" Then
            ' Player one
            piece1 = "B1"
            piece2 = "G1"
            possibleNextPositionP1_1 = Array("A3", "C3")
            possibleNextPositionP1_2 = Array("F3", "H3")
            
            ' Player two
            piece3 = "B8"
            piece4 = "G8"
            possibleNextPositionP2_1 = Array("A6", "C6")
            possibleNextPositionP2_2 = Array("F6", "H6")
        ElseIf piece = "Bishop" Then
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
        pieceP1_1.Add "nextPos", possibleNextPositionP1_1
        pieceP1_2.Add "nextPos", possibleNextPositionP1_2
        pieceP1_1.Add "type", piece
        pieceP1_2.Add "type", piece
        pieceP1_1.Add "piecesEater", Empty
        pieceP1_2.Add "piecesEater", Empty
        
        ' Player two
        pieceP2_1.Add "firstPos", piece3
        pieceP2_2.Add "firstPos", piece4
        pieceP2_1.Add "newPos", piece3
        pieceP2_2.Add "newPos", piece4
        pieceP2_1.Add "nextPos", possibleNextPositionP2_1
        pieceP2_2.Add "nextPos", possibleNextPositionP2_2
        pieceP2_1.Add "type", piece
        pieceP2_2.Add "type", piece
        pieceP2_1.Add "piecesEater", Empty
        pieceP2_2.Add "piecesEater", Empty
        

        Dim valueIsFirstMove as boolean
        For Each value In Array("moved", "firstMove", "dead")
            valueIsFirstMove = iif(CStr(value) = "firstMove", True, False)
            pieceP1_1.Add CStr(value), valueIsFirstMove
            pieceP1_2.Add CStr(value), valueIsFirstMove
            pieceP2_1.Add CStr(value), valueIsFirstMove
            pieceP2_2.Add CStr(value), valueIsFirstMove
        Next value
        
        i = 1
        For Each value In Array(piece1, piece2, piece3, piece4)
            buttons(CStr(value))("isPiece") = True
            buttons(CStr(value))("player") = iif(i > 2, 2, 1)
            buttons(CStr(value))("piece") = CStr(value) & piece
            i = i + 1
        Next value
        
        playerOne.Add piece1 & piece, pieceP1_1
        playerOne.Add piece2 & piece, pieceP1_2
        
        playerTwo.Add piece3 & piece, pieceP2_1
        playerTwo.Add piece4 & piece, pieceP2_2
        
    Next piece
    
    ' Queen and King -----------------------------------------------------------
    For Each piece In Array("Queen", "King")
        Set chessPiece1 = CreateObject("Scripting.Dictionary")
        Set chessPiece2 = CreateObject("Scripting.Dictionary")
        
        piece1 = iif(piece = "Queen", "D1", "E1")
        piece2 = iif(piece = "Queen", "D8", "E8")
        
        For Each value In Array(piece1, piece2)
            buttons(value)("isPiece") = True
            buttons(value)("player") = iif(value = piece1, 1, 2)
            buttons(value)("piece") = value & piece
        Next value
        
        chessPiece1.Add "firstPos", piece1
        chessPiece2.Add "firstPos", piece2
        chessPiece1.Add "newPos", piece1
        chessPiece2.Add "newPos", piece2
        chessPiece1.Add "nextPos", Empty
        chessPiece2.Add "nextPos", Empty
        chessPiece1.Add "type", piece
        chessPiece2.Add "type", piece
        
        For Each value In Array("firstMove", "moved", "danger", "dead")
            chessPiece1.Add value, iif(value = "firstMove", True, False)
            chessPiece2.Add value, iif(value = "firstMove", True, False)
        Next value

        chessPiece1.Add "piecesEater", Empty
        chessPiece2.Add "piecesEater", Empty
        
        
        playerOne.Add piece1 & piece, chessPiece1
        playerTwo.Add piece2 & piece, chessPiece2
        
    Next piece
    
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
    
    For Each value In colors.keys
        Controls("L" & CStr(value)).BackColor = colors(value)
    Next value
    rePaintCases
    
End Function