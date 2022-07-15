SText = InputBox("TEXT","Spamer")
Oft = InputBox("How many times","Spamer")

If MsgBox("Click on Chat Field.." & vbNewLine & "Then Yes to send..", vbYesNo + vbQuestion + vbSystemModal, "Spamer") = vbYes Then

Set SpamerWin = WScript.CreateObject("WScript.Shell")

WScript.Sleep 1000

For i = 1 to oft
WScript.Sleep 10
SpamerWin.SendKeys SText
WScript.Sleep 10
SpamerWin.SendKeys "{ENTER}"
Next
'SpamerWin.SendKeys "{ENTER}"
WScript.Sleep 2000
MsgBox "Finished...", 1024 + vbSystemModal, "Spamer"

Else
MsgBox "Canceled", vbSystemModal, "Spamer"
End If