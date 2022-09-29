Dim a : a = 10
Dim b : b = 0
Dim c

If a <> 0 AND b <> 0 Then
  x = msgbox ("AND Resultado do operado é : True")
Else
  x = msgbox ("AND Resultado do operado é : False")
End If

If a <> 0 OR b <> 0 Then
  x = msgbox ("OR Resultado do operado é : True")
Else
  x = msgbox ("OR Resultado do operado é : False")
End If

If NOT (a <> 0 OR b <> 0) Then
  x = msgbox ("NOT Resultado do operado é : True")
Else
  x = msgbox ("NOT Resultado do operado é : False")
End If

If (a <> 0 XOR b <> 0) Then
  x=msgbox ("XOR Resultado do operado é : True")
Else
  x=msgbox ("XOR Resultado do operado é : False")
End If