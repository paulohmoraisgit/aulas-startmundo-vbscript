Dim input
input = InputBox("Qual seu nome?", "Preencha")

If input = "" Then
  MsgBox "Digite algo."
Else
  MsgBox "O nome digitado foi " & input
End If 