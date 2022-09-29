SET wsite = CreateObject("WScript.Shell")

website = InputBox("Meu buscador", "Buscador")

If website = "" Then
  MsgBox "Para que seja feita uma busca, eh preciso digitar algo!"
Else
  wsite.Run "https://www.google.com/search?q=" & website
End If