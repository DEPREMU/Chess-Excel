Public Function getAvailablePosP1(piece As String, Optional emulatePiece As Variant) As Variant
    getAvailablePosP1 = Empty

    If playerOne(piece)("dead") Then Exit Function

    If IsMissing(emulatePiece) Then emulatePiece = Array("", "")

    Dim pieceType As String
    Dim indexLetter As Integer
    Dim number As String
    Dim letter As String
    Dim btn As String
    Dim valueToReturn As Object
    Dim btnLocal As String
    Dim i As Integer
    Dim j As Integer
    Dim newNum As Integer
    Dim player As Integer
    Dim availablePos As Object
    Dim value As Variant
    Dim leftOrRight As String
    Dim posP2 As String

    Set availablePos = CreateObject("Scripting.Dictionary")
    Set valueToReturn = CreateObject("Scripting.Dictionary")

    pieceType = playerOne(piece)("type")

    Select Case pieceType
        Case "Pawn"
            values = NextPosPawn(piece, True)
            If Not IsEmpty(values) Then getAvailablePosP1 = values
            Set availablePos = Nothing
            Set valueToReturn = Nothing
            Exit Function
        Case "Rook"
            values = getNextPosRook(piece, True, emulatePiece)
            If Not IsEmpty(values) Then getAvailablePosP1 = values
            Set availablePos = Nothing
            Set valueToReturn = Nothing
            Exit Function

        Case "Knight"
            getAvailablePosP1 = getNextPosKnight(piece, True)
            Set availablePos = Nothing
            Set valueToReturn = Nothing
            Exit Function

        Case "Bishop"
            getAvailablePosP1 = getNextPosBishop(piece, True, emulatePiece)
            Set availablePos = Nothing
            Set valueToReturn = Nothing
            Exit Function

        Case "Queen"
            values = getNextPosRook(piece, True, emulatePiece)
            If Not IsEmpty(values) Then
                For Each value In values
                    availablePos.Add value, True
                Next value
            End If

            values = getNextPosBishop(piece, True, emulatePiece)
            If Not IsEmpty(values) Then
                For Each value In values
                    availablePos.Add value, True
                Next value
            End If

        Case "King"
            values = getPosKingP1(piece, CStr(playerOne(piece)("newPos")))
            If Not IsEmpty(values) Then getAvailablePosP1 = values
            Set availablePos = Nothing
            Set valueToReturn = Nothing
            Exit Function
    End Select

    i = 1
    For Each value In availablePos.keys
        If availablePos(value) Then
            valueToReturn.Add CStr(i), value
            i = i + 1
        End If
    Next value
    Set availablePos = Nothing
    If i = 1 Then
        Set valueToReturn = Nothing
        Exit Function
    End If

    getAvailablePosP1 = valueToReturn.items
    
    Set valueToReturn = Nothing
End Function


Public Function getPosKingP1(piece As String, position As String) As Variant
    Dim letter As String
    Dim indexLetter As Integer
    Dim number As String
    Dim valuesAdded As Integer
    Dim availablePos As Object
    Set availablePos = CreateObject("Scripting.Dictionary")

    letter = Mid(position, 1, 1)
    indexLetter = numbers(letter)
    number = Mid(position, 2, 1)
    valuesAdded = 0

    Dim directions As Variant
    directions = Array(Array(0, 1), Array( - 1, 1), Array(1, 1), Array( - 1, 0), Array(1, 0), Array(0, - 1), Array( - 1, - 1), Array(1, - 1))

    For i = LBound(directions) To UBound(directions)

        newLetterIndex = indexLetter + directions(i)(0)
        newNumber = CInt(number) + directions(i)(1)

        If newLetterIndex >= 1 And newLetterIndex <= 8 And newNumber >= 1 And newNumber <= 8 Then
            btn = letters(CStr(newLetterIndex)) & CStr(newNumber)
            If Not buttons(btn)("isPiece") Or buttons(btn)("player") = 2 And btn <> playerOne(piece)("newPos") Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
        End If
    Next i

    If playerOne(piece)("firstMove") Then
        ' Enroque corto
        If Not buttons("F1")("isPiece") And Not buttons("G1")("isPiece") And playerOne("H1Rook")("firstMove") And Not playerOne("H1Rook")("dead") Then
            availablePos.Add "G1", True
            valuesAdded = valuesAdded + 1
        End If

        ' Enroque largo
        If Not buttons("B1")("isPiece") And Not buttons("C1")("isPiece") And Not buttons("D1")("isPiece") And playerOne("A1Rook")("firstMove") And Not playerOne("A1Rook")("dead") Then
            availablePos.Add "C1", True
            valuesAdded = valuesAdded + 1
        End If
    End If
    If valuesAdded = 0 Then
        getPosKingP1 = Empty
    Else
        getPosKingP1 = availablePos.keys
    End If
    
    Set availablePos = Nothing
End Function