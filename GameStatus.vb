Option Explicit


Dim frm As UserForm


Private Sub Init()
    Set frm = UserForms(0)
End Sub


Public Function checkGameStatus(piece As String)
    If frm Is Nothing Then Init

    If isCheckMate("E1King", True) Then
        MsgBox "Player two wins"
        frm.Comments.Caption = "W Player Two"
    ElseIf isCheckMate("E8King", False) Then
        MsgBox "Player one wins"
        frm.Comments.Caption = "W Player One"
    End If
End Function


Public Function isCheck(piece As String, position As String, boolPlayerOne As Boolean)
    Dim pieceP1 As Variant
    Dim pieceP2 As Variant
    Dim availablePosP1 As Variant
    Dim availablePosP2 As Variant
    
    
    If boolPlayerOne Then
        For Each pieceP2 In playerTwo.keys
            If playerTwo(pieceP2)("dead") Then Goto ContinueLoopP1
            availablePosP2 = getAvailablePosP2(CStr(pieceP2))
            If IsEmpty(availablePosP2) Then Goto ContinueLoopP1
            
            If ArrayContains(availablePosP2, position) Then
                isCheck = True
                Exit Function
            End If
            
            ContinueLoopP1 :
        Next pieceP2
        
    Else
        For Each pieceP1 In playerOne.keys
            If playerOne(pieceP1)("dead") Then Goto ContinueLoopP2
            availablePosP1 = getAvailablePosP1(CStr(pieceP1))
            If IsEmpty(availablePosP1) Then Goto ContinueLoopP2
            
            If ArrayContains(availablePosP1, position) Then
                isCheck = True
                Exit Function
            End If
            
            ContinueLoopP2 :
        Next pieceP1
        
    End If
    isCheck = False
End Function


Public Function isCheckMate(piece As String, boolPlayerOne As Boolean) As Boolean
    Dim availablePos As Variant
    Dim countPos As Integer
    Dim countMate As Integer
    Dim pos As Variant
    countMate = 0
    countPos = 0

    If boolPlayerOne Then
        availablePos = possiblePosKingP1(piece, CStr(playerOne(piece)("newPos")))
        If IsEmpty(availablePos) Then
            isCheckMate = False
            Exit Function
        End If
        For Each pos In availablePos
            If isCheck(piece, CStr(pos), True) Then countMate = countMate + 1
            countPos = countPos + 1
        Next pos
        If countPos = countMate Then
            isCheckMate = True
            Exit Function
        End If
    Else
        availablePos = posiblePosKingP2(piece, CStr(playerTwo(piece)("newPos")))
        If IsEmpty(availablePos) Then
            isCheckMate = False
            Exit Function
        End If
        For Each pos In availablePos
            If isCheck(piece, CStr(pos), False) Then countMate = countMate + 1
            countPos = countPos + 1
        Next pos
        If countPos = countMate Then
            isCheckMate = True
            Exit Function
        End If
    End If
    isCheckMate = False
End Function