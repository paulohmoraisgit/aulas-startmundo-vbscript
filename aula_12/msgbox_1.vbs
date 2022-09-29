Set network = WScript.CreateObject("WScript.Network")

info = "My Domain: " & vbTab & network.UserDomain & vbCrLf & "My computer: " & vbTab & network.ComputerName & vbCrLf & "My user: " & vbTab & network.UserName & vbCrLf

MsgBox info, vbOKOnly, "System and network information"