Option Explicit

Public Function getNextPos(piece as String, optional emulatePiece as variant) as Variant

    Dim boolPlayerOne as Boolean
    
    If IsMissing(emulatePiece) Then emulatePiece = Array("", "")

    If isPlayerOnePiece(piece) Then
        boolPlayerOne = True
    Else
        boolPlayerOne = False
    End If
    pieceType = getTypePiece(piece, boolPlayerOne)

    Select Case pieceType
        Case "Pawn"
            getNextPos = getNextPosPawn(piece, boolPlayerOne)
        Case "Rook"
            getNextPos = getNextPosRook(piece, boolPlayerOne, emulatePiece)
        Case "Knight"
            getNextPos = getNextPosKnight(piece, boolPlayerOne)
        Case "Bishop"
            getNextPos = getNextPosBishop(piece, boolPlayerOne, emulatePiece)
        Case "Queen"
            getNextPos = mergeArrs( _
                getNextPosRook(piece, boolPlayerOne, emulatePiece), _
                getNextPosBishop(piece, boolPlayerOne, emulatePiece) _
                )
        Case "King"
            getNextPos = getNextPosKing(piece, CStr(getPosPlayer(piece, boolPlayerOne)), boolPlayerOne)
    End Select
End Function