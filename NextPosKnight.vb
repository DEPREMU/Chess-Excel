Option Explicit

Public Function getNextPosKnight(piece As String, boolPlayerOne As Boolean) As Variant
    Dim newLetterIndex As Integer
    Dim offsets As Variant
    Dim position As String
    Dim letter As String
    Dim indexLetter As Integer
    Dim number As Integer
    Dim newNum As Integer
    Dim btn As String
    Dim player As Integer
    Dim i As Integer
    Dim availablePos As Object
    Dim valuesAdded As Integer
    Dim localPiece As String

    Set availablePos = CreateObject("Scripting.Dictionary")


    position = getPosPlayer(piece, boolPlayerOne)
    letter = Mid(position, 1, 1)
    indexLetter = numbers(letter)
    number = CInt(Mid(position, 2, 1))

    offsets = Array(Array(2, -1), Array(2, 1), Array(-2, -1), Array(-2, 1), Array(1, -2), Array(1, 2), Array(-1, -2), Array(-1, 2))

    For i = LBound(offsets) To UBound(offsets)
        newNum = number + offsets(i)(0)
        newLetterIndex = indexLetter + offsets(i)(1)

        If newNum >= 1 And newNum <= 8 And newLetterIndex >= 1 And newLetterIndex <= 8 Then
            btn = letters(CStr(newLetterIndex)) & CStr(newNum)
            player = buttons(btn)("player")
            localPiece = buttons(btn)("piece")

            If player = 0 Then
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            ElseIf player = 1 And Not boolPlayerOne Then
                playerOne(localPiece)("danger") = True
                playerOne(localPiece)("piecesEater") = addToArr(playerOne(localPiece)("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            ElseIf player = 2 And boolPlayerOne Then
                playerTwo(localPiece)("danger") = True
                playerTwo(localPiece)("piecesEater") = addToArr(playerTwo(localPiece)("piecesEater"), piece)
                availablePos.Add btn, True
                valuesAdded = valuesAdded + 1
            End If
        End If
    Next i

    If valuesAdded = 0 Then Exit Function
    
    getNextPosKnight = availablePos.keys
    Set availablePos = Nothing
End Function
