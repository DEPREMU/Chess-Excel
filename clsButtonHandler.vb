Public WithEvents btn As MSForms.CommandButton
Public WithEvents label As MSForms.label

Private Sub btn_Click()
    If Not ArrayContains(Array("ButtonStartGame", "ButtonRestartGame"), btn.name) Then
        Call GenericClick(btn)
    End If
End Sub

Private Sub label_Click()
    If Mid(label.name, 1, 1) <> "L" Then
        Call GenericClickLabel(label)
    End If
End Sub