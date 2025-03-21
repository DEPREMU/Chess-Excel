Public Function getAvailablePosP2(piece As String, Optional emulatePiece As Variant) As Variant
	If playerTwo(piece)("dead") Then
		getAvailablePosP2 = Empty
		Exit Function
	End If
	
	If IsMissing(emulatePiece) Then emulatePiece = Array("", "")
	
	Dim pieceType As String
	Dim indexLetter As Integer
	Dim number As String
	Dim letter As String
	Dim btn As String
	Dim valueToReturn As Object
	Dim i As Integer
	Dim j As Integer
	Dim newNum As Integer
	Dim player As Integer
	Dim availablePos As Object
	Dim value As Variant
	Dim posP1 As String
	Dim leftOrRight As Integer
	
	Set availablePos = CreateObject("Scripting.Dictionary")
	Set valueToReturn = CreateObject("Scripting.Dictionary")
	
	pieceType = playerTwo(piece)("type")
	
	Select Case pieceType
			
		Case "Pawn"
			letter = Mid(playerTwo(piece)("newPos"), 1, 1)
			indexLetter = numbers(letter)
			number = Mid(playerTwo(piece)("newPos"), 2, 1)
			
			For i = 0 To 1
				leftOrRight = IIf(i = 0, 1, - 1)
				If letters(CStr(indexLetter - leftOrRight)) <> "" Then
					btn = letters(CStr(indexLetter - leftOrRight)) & number
					If btn <> "" Then
						If buttons(btn)("player") = 1 Then
							If playerOne(buttons(btn)("piece"))("enPassant") Then
								posP1 = playerOne(buttons(btn)("piece"))("newPos")
								availablePos.Add Mid(posP1, 1, 1) & CStr(CInt(Mid(posP1, 2, 1)) - 1), True
							End If
						End If
					End If
				End If
			Next i
			
			' Move forward
			If CInt(number) > 1 Then
				btn = letter & CStr(CInt(number) - 1)
				If Not buttons(btn)("isPiece") Then
					availablePos.Add btn, True
					If playerTwo(piece)("firstMove") Then
						If Not buttons(letter & CStr(CInt(number) - 2))("isPiece") Then availablePos.Add letter & CStr(CInt(number) - 2), True
					End If
				End If
			End If
			
			' Capture diagonally left
			If indexLetter > 1 And CInt(number) > 1 Then
				btn = letters(CStr(indexLetter - 1)) & CStr(CInt(number) - 1)
				If buttons(btn)("player") = 1 Then
					availablePos.Add btn, True
					playerOne(buttons(btn)("piece"))("danger") = True
					playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
				End If
			End If
			
			' Capture diagonally right
			If indexLetter < 8 And CInt(number) > 1 Then
				btn = letters(CStr(indexLetter + 1)) & CStr(CInt(number) - 1)
				If buttons(btn)("player") = 1 Then
					playerOne(buttons(btn)("piece"))("danger") = True
					playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
					availablePos.Add btn, True
				End If
			End If
			
		Case "Rook"
			values = getAvailablePosRookP2(piece, emulatePiece)
			If IsEmpty(values) Then
				getAvailablePosP2 = Empty
			Else
				getAvailablePosP2 = values
			End If
			Exit Function
			
		Case "Knight"
			Dim newLetterIndex As Integer
			Dim offsets As Variant
			Dim position As String
			
			position = playerTwo(piece)("newPos")
			letter = Mid(position, 1, 1)
			indexLetter = numbers(letter)
			number = CInt(Mid(position, 2, 1))
			
			offsets = Array(Array(2, - 1), Array(2, 1), Array( - 2, - 1), Array( - 2, 1), Array(1, - 2), Array(1, 2), Array( - 1, - 2), Array( - 1, 2))
			
			For i = LBound(offsets) To UBound(offsets)
				newNum = number + offsets(i)(0)
				newLetterIndex = indexLetter + offsets(i)(1)
				
				If newNum >= 1 And newNum <= 8 And newLetterIndex >= 1 And newLetterIndex <= 8 Then
					btn = letters(CStr(newLetterIndex)) & CStr(newNum)
					player = buttons(btn)("player")
					
					If player = 1 Or player = 0 Then
						If player = 1 Then
							playerOne(buttons(btn)("piece"))("danger") = True
							playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
						End If
						availablePos.Add btn, True
					End If
				End If
			Next
			
		Case "Bishop"
			values = getAvailablePosBishopP2(piece, emulatePiece)
			If Not IsEmpty(values) Then
				getAvailablePosP2 = values
			Else
				getAvailablePosP2 = Empty
			End If
			Exit Function
			
		Case "Queen"
			values = getAvailablePosRookP2(piece, emulatePiece)
			If Not IsEmpty(values) Then
				For Each value In values
					availablePos.Add value, True
				Next value
			End If
			
			values = getAvailablePosBishopP2(piece, emulatePiece)
			If Not IsEmpty(values) Then
				For Each value In values
					availablePos.Add value, True
				Next value
			End If
			
		Case "King"
			values = posiblePosKingP2(piece, CStr(playerTwo(piece)("newPos")))
			If Not IsEmpty(values) Then
				getAvailablePosP2 = values
			Else
				getAvailablePosP2 = Empty
			End If
			Exit Function
			
	End Select
	
	i = 1
	For Each value In availablePos.keys
		If availablePos(value) Then
			valueToReturn.Add CStr(i), value
			i = i + 1
		End If
	Next value
	If i = 1 Then
		getAvailablePosP2 = Empty
		Exit Function
	End If
	getAvailablePosP2 = valueToReturn.items
End Function

Public Function posiblePosKingP2(piece As String, position As String) As Variant
	
	Dim letter As String
	Dim indexLetter As Integer
	Dim number As String
	Dim valuesAdded As Integer
	Dim availablePos As Object
	Set availablePos = CreateObject("Scripting.Dictionary")
	valuesAdded = 0
	
	letter = Mid(playerTwo(piece)("newPos"), 1, 1)
	indexLetter = numbers(letter)
	number = Mid(playerTwo(piece)("newPos"), 2, 1)
	
	Dim directions As Variant
	directions = Array(Array(0, 1), Array( - 1, 1), Array(1, 1), Array( - 1, 0), Array(1, 0), Array(0, - 1), Array( - 1, - 1), Array(1, - 1))
	
	For i = LBound(directions) To UBound(directions)
		
		newLetterIndex = indexLetter + directions(i)(0)
		newNumber = CInt(number) + directions(i)(1)
		
		If newLetterIndex >= 1 And newLetterIndex <= 8 And newNumber >= 1 And newNumber <= 8 Then
			btn = letters(CStr(newLetterIndex)) & CStr(newNumber)
			If Not buttons(btn)("isPiece") Or buttons(btn)("player") = 1 And btn <> playerTwo(piece)("newPos") Then
				availablePos.Add btn, True
				valuesAdded = valuesAdded + 1
			End If
		End If
	Next i
	
	If playerTwo(piece)("firstMove") Then
		' Enroque corto
		If Not buttons("F8")("isPiece") And Not buttons("G8")("isPiece") And playerTwo("H8Rook")("firstMove") And Not playerTwo("H8Rook")("dead") Then
			availablePos.Add "G8", True
			valuesAdded = valuesAdded + 1
		End If
		
		' Enroque largo
		If Not buttons("B8")("isPiece") And Not buttons("C8")("isPiece") And Not buttons("D8")("isPiece") And playerTwo("A8Rook")("firstMove") And Not playerTwo("A8Rook")("dead") Then
			availablePos.Add "C8", True
			valuesAdded = valuesAdded + 1
		End If
	End If
	
	If valuesAdded = 0 Then
		posiblePosKingP2 = Empty
	Else
		posiblePosKingP2 = availablePos.keys
	End If
End Function

Public Function getAvailablePosRookP2(piece As String, emulatePiece As Variant) As Variant
	Dim letter As String
	Dim indexLetter As Integer
	Dim number As String
	Dim i As Integer
	Dim j As Integer
	Dim btn As String
	Dim availablePos As Object
	Dim value As Variant
	Dim valuesAdded As Integer
	Dim lastIsPieceBtns As Boolean
	Dim lastPos As Variant
	Dim lastPlayer As Integer
	Dim lastPiece As Variant
	Dim lastIsPiece As Boolean
	valuesAdded = 0
	
	If IsMissing(emulatePiece) Then emulatePiece = Array("", "")
	If emulatePiece(0) <> "" And emulatePiece(1) <> "" Then
		lastPos = playerOne(emulatePiece(0))("newPos")
		lastPlayer = buttons(playerOne(emulatePiece(0))("newPos"))("player")
		lastPiece = buttons(playerOne(emulatePiece(0))("newPos"))("piece")
		lastIsPiece = buttons(playerOne(emulatePiece(0))("newPos"))("isPiece")
		lastIsPieceBtns = buttons(emulatePiece(1))("isPiece")
		
		buttons(playerOne(emulatePiece(0))("newPos"))("isPiece") = False
		buttons(playerOne(emulatePiece(0))("newPos"))("player") = 0
		playerOne(emulatePiece(0))("newPos") = emulatePiece(1)
		buttons(emulatePiece(1))("isPiece") = True
	End If
	
	Set availablePos = CreateObject("Scripting.Dictionary")
	
	letter = Mid(playerTwo(piece)("newPos"), 1, 1)
	indexLetter = numbers(letter)
	number = Mid(playerTwo(piece)("newPos"), 2, 1)
	
	' Top
	For i = CInt(number) + 1 To 8
		btn = letter & CStr(i)
		If emulatePiece(1) = btn Then
			availablePos.Add btn, True
			valuesAdded = valuesAdded + 1
			playerOne(emulatePiece(0))("danger") = True
			playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
			Exit For
		ElseIf buttons(btn)("isPiece") Then
			If buttons(btn)("player") = 1 Then
				availablePos.Add btn, True
				valuesAdded = valuesAdded + 1
				playerOne(buttons(btn)("piece"))("danger") = True
				playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
			End If
			Exit For
		End If
		valuesAdded = valuesAdded + 1
		availablePos.Add btn, True
	Next i
	
	' Bottom
	For Each value In range(CInt(number), 0, - 1)
		If number <> value And value Then
			btn = letter & CStr(value)
			If emulatePiece(1) = btn Then
				availablePos.Add btn, True
				valuesAdded = valuesAdded + 1
				playerOne(emulatePiece(0))("danger") = True
				playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
				Exit For
			ElseIf buttons(btn)("isPiece") Then
				If buttons(btn)("player") = 1 Then
					availablePos.Add btn, True
					valuesAdded = valuesAdded + 1
					playerOne(buttons(btn)("piece"))("danger") = True
					playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
				End If
				Exit For
			End If
			availablePos.Add btn, True
			valuesAdded = valuesAdded + 1
		End If
	Next value
	
	' Left
	For Each value In Array("H", "G", "F", "E", "D", "C", "B", "A")
		If indexLetter = 1 Then Exit For
		If indexLetter < numbers(value) Or value = letter Then Goto ContinueLoop
		
		btn = value & number
		If emulatePiece(1) = btn Then
			playerOne(emulatePiece(0))("danger") = True
			playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
			availablePos.Add btn, True
			valuesAdded = valuesAdded + 1
			Exit For
		ElseIf buttons(btn)("isPiece") Then
			If buttons(btn)("player") = 1 Then
				playerOne(buttons(btn)("piece"))("danger") = True
				playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
				availablePos.Add btn, True
				valuesAdded = valuesAdded + 1
			End If
			Exit For
		End If
		availablePos.Add btn, True
		valuesAdded = valuesAdded + 1
		
		Goto ContinueLoop
		ContinueLoop :
	Next value
	
	' Right
	For Each value In Array("A", "B", "C", "D", "E", "F", "G", "H")
		If indexLetter = 8 Then Exit For
		If numbers(value) > indexLetter Then
			btn = value & number
			If emulatePiece(1) = btn Then
				playerOne(emulatePiece(0))("danger") = True
				playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
				availablePos.Add btn, True
				valuesAdded = valuesAdded + 1
				Exit For
			ElseIf buttons(btn)("isPiece") Then
				If buttons(btn)("player") = 1 Then
					playerOne(buttons(btn)("piece"))("danger") = True
					playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
					availablePos.Add btn, True
					valuesAdded = valuesAdded + 1
				End If
				Exit For
			End If
			availablePos.Add btn, True
			valuesAdded = valuesAdded + 1
		End If
	Next value
	
	If emulatePiece(0) <> "" And emulatePiece(1) <> "" Then
		playerOne(emulatePiece(0))("newPos") = lastPos
		buttons(playerOne(emulatePiece(0))("newPos"))("isPiece") = lastIsPiece
		buttons(playerOne(emulatePiece(0))("newPos"))("player") = lastPlayer
		buttons(playerOne(emulatePiece(0))("newPos"))("piece") = lastPiece
		buttons(emulatePiece(1))("isPiece") = lastIsPieceBtns
	End If
	
	If valuesAdded = 0 Then
		getAvailablePosRookP2 = Empty
	Else
		getAvailablePosRookP2 = availablePos.keys
	End If
End Function

Public Function getAvailablePosBishopP2(piece As String, emulatePiece As Variant) As Variant
	Dim letter As String
	Dim indexLetter As Integer
	Dim number As String
	Dim i As Integer
	Dim j As Integer
	Dim btn As String
	Dim availablePos As Object
	Dim value As Variant
	Dim valuesAdded As Integer
	Dim lastIsPieceBtns As Boolean
	Dim lastPos As Variant
	Dim lastPlayer As Integer
	Dim lastPiece As Variant
	Dim lastIsPiece As Boolean
	valuesAdded = 0
	
	If IsMissing(emulatePiece) Then emulatePiece = Array("", "")
	If emulatePiece(0) <> "" And emulatePiece(1) <> "" Then
		lastPos = playerOne(emulatePiece(0))("newPos")
		lastPlayer = buttons(playerOne(emulatePiece(0))("newPos"))("player")
		lastPiece = buttons(playerOne(emulatePiece(0))("newPos"))("piece")
		lastIsPiece = buttons(playerOne(emulatePiece(0))("newPos"))("isPiece")
		lastIsPieceBtns = buttons(emulatePiece(1))("isPiece")
		
		buttons(playerOne(emulatePiece(0))("newPos"))("isPiece") = False
		buttons(playerOne(emulatePiece(0))("newPos"))("player") = 0
		playerOne(emulatePiece(0))("newPos") = emulatePiece(1)
		buttons(emulatePiece(1))("isPiece") = True
	End If
	
	Set availablePos = CreateObject("Scripting.Dictionary")
	
	letter = Mid(playerTwo(piece)("newPos"), 1, 1)
	indexLetter = numbers(letter)
	number = Mid(playerTwo(piece)("newPos"), 2, 1)
	
	' Top Left
	i = indexLetter
	j = CInt(number)
	Do While i > 1 And j < 8
		i = i - 1
		j = j + 1
		btn = letters(CStr(i)) & CStr(j)
		If emulatePiece(1) = btn Then
			playerOne(emulatePiece(0))("danger") = True
			playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
			availablePos.Add btn, True
			valuesAdded = valuesAdded + 1
			Exit Do
		ElseIf buttons(btn)("isPiece") Then
			If buttons(btn)("player") = 1 Then
				playerOne(buttons(btn)("piece"))("danger") = True
				playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
				availablePos.Add btn, True
				valuesAdded = valuesAdded + 1
			End If
			Exit Do
		End If
		availablePos.Add btn, True
		valuesAdded = valuesAdded + 1
	Loop
	
	' Top Right
	i = indexLetter
	j = CInt(number)
	Do While i < 8 And j < 8
		i = i + 1
		j = j + 1
		btn = letters(CStr(i)) & CStr(j)
		If emulatePiece(1) = btn Then
			playerOne(emulatePiece(0))("danger") = True
			playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
			availablePos.Add btn, True
			valuesAdded = valuesAdded + 1
			Exit Do
		ElseIf buttons(btn)("isPiece") Then
			If buttons(btn)("player") = 1 Then
				playerOne(buttons(btn)("piece"))("danger") = True
				playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
				availablePos.Add btn, True
				valuesAdded = valuesAdded + 1
			End If
			Exit Do
		End If
		availablePos.Add btn, True
		valuesAdded = valuesAdded + 1
	Loop
	
	' Bottom Left
	i = indexLetter
	j = CInt(number)
	Do While i > 1 And j > 1
		i = i - 1
		j = j - 1
		btn = letters(CStr(i)) & CStr(j)
		If emulatePiece(1) = btn Then
			playerOne(emulatePiece(0))("danger") = True
			playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
			availablePos.Add btn, True
			valuesAdded = valuesAdded + 1
			Exit Do
		ElseIf buttons(btn)("isPiece") Then
			If buttons(btn)("player") = 1 Then
				playerOne(buttons(btn)("piece"))("danger") = True
				playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
				availablePos.Add btn, True
				valuesAdded = valuesAdded + 1
			End If
			Exit Do
		End If
		availablePos.Add btn, True
		valuesAdded = valuesAdded + 1
	Loop
	
	' Bottom Right
	i = indexLetter
	j = CInt(number)
	Do While i < 8 And j > 1
		i = i + 1
		j = j - 1
		btn = letters(CStr(i)) & CStr(j)
		If emulatePiece(1) = btn Then
			playerOne(emulatePiece(0))("danger") = True
			playerOne(emulatePiece(0))("piecesEater") = addToArr(playerOne(emulatePiece(0))("piecesEater"), piece)
			availablePos.Add btn, True
			valuesAdded = valuesAdded + 1
			Exit Do
		ElseIf buttons(btn)("isPiece") Then
			If buttons(btn)("player") = 1 Then
				playerOne(buttons(btn)("piece"))("danger") = True
				playerOne(buttons(btn)("piece"))("piecesEater") = addToArr(playerOne(buttons(btn)("piece"))("piecesEater"), piece)
				availablePos.Add btn, True
				valuesAdded = valuesAdded + 1
			End If
			Exit Do
		End If
		availablePos.Add btn, True
		valuesAdded = valuesAdded + 1
	Loop

	If emulatePiece(0) <> "" And emulatePiece(1) <> "" Then
		playerOne(emulatePiece(0))("newPos") = lastPos
		buttons(playerOne(emulatePiece(0))("newPos"))("isPiece") = lastIsPiece
		buttons(playerOne(emulatePiece(0))("newPos"))("player") = lastPlayer
		buttons(playerOne(emulatePiece(0))("newPos"))("piece") = lastPiece
		buttons(emulatePiece(1))("isPiece") = lastIsPieceBtns
	End If
	
	If valuesAdded = 0 Then
		getAvailablePosBishopP2 = Empty
	Else
		getAvailablePosBishopP2 = availablePos.keys
	End If
End Function