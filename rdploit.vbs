Dim objShell, strComputer, strUser, intWidth, intHeight

Set objShell = CreateObject("WScript.Shell")

' 1. Prompt for Remote Computer Name or IP
strComputer = InputBox("Enter Remote Computer Name or IP Address:", "RDP Connection Manager")

' Exit if user cancels or leaves empty
If Trim(strComputer) = "" Then WScript.Quit

' 2. Prompt for Username (Optional)
strUser = InputBox("Enter Username (leave blank if not needed):", "RDP Connection Manager", "Administrator")

' 3. Formulate and launch the MSTSC command
If Trim(strUser) = "" Then
    ' Launch with just the computer name
    objShell.Run "mstsc /v:" & strComputer
Else
    ' Launch with the computer name and pre-fill username
    ' Note: For security reasons, Microsoft does not allow passing passwords via command line arguments.
    objShell.Run "mstsc /v:" & strComputer & " /user:" & strUser
End If

Set objShell = Nothing
