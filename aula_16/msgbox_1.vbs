Set WshShell = WScript.CreateObject("WScript.Shell")
Dim opcao

opcao = MsgBox("Escolha a opcao", vbYesNo)

If opcao = vbYes Then
  WshShell.run "Calc.exe"
Else
  WshShell.run "notepad"
End If