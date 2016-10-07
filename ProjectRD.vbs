Function createIT() 'Function for the IT account creator
	Set shell = CreateObject("Shell.Application")
	shell.ShellExecute "cmd.exe", "/c net user /add IT Lbsbgq33# & net localgroup administrators IT /add", "", "runas", 1 
	Wscript.echo "done"
end Function 'will need to set a password on this account at the end of the program'

Function otherUserRun()'function obsolete due to windows 7 compatibility issues'
	Set shell = CreateObject("WScript.Shell")
	set fso = CreateObject("Scripting.FileSystemObject")
	If(fso.FolderExists("C:\Program Files (x86)")) Then 'This if statement checks for the existence of the (x86) folder'
		Wscript.echo "found it"
		shell.run "cmd /k runas /user:IT ""C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
		WScript.Sleep 1000 'just making sure the password prompt is loaded'
		shell.SendKeys "Lbsbgq33#"
		shell.SendKeys "~+exit+~"
	Else
		Wscript.echo "did not find it"
		shell.run "cmd /k runas /user:IT ""C:\Program Files\Google\Chrome\Application\chrome.exe"""
		WScript.Sleep 1000 'just making sure the password prompt is loaded'
		shell.SendKeys "Lbsbgq33#"
		shell.SendKeys "~+exit+~"
	End If
	Wscript.Sleep 5000 'making sure chrome is open and running'
	temp = shell.AppActivate("Chrome")
	If(temp = -1) Then
		shell.SendKeys "^+m"
		WScript.Sleep 2000
		shell.SendKeys "~"
		WScript.Sleep 2000
		shell.SendKeys "%{F4}"
		WScript.Sleep 2000
		shell.SendKeys "zac.adams@versacor.com"
		WScript.Sleep 2000
		shell.SendKeys "~"
		WScript.Sleep 1000
		shell.SendKeys "Lbsbgq33#"
		WScript.Sleep 1000
		shell.SendKeys "~"
		WScript.Sleep 3000
		shell.SendKeys "{TAB}"
		WScript.Sleep 1000
		shell.SendKeys "{TAB}"
		WScript.Sleep 1000
		shell.SendKeys "~"
		WScript.Sleep 1000
		shell.run "cmd /k runas /user:IT ""C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
		WScript.Sleep 2000
		shell.SendKeys "Lbsbgq33#"
		WScript.Sleep 20000
		shell.run "cmd /k runas /user:IT ""C:\Program Files (x86)\Google\Chrome\Application\chrome.exe  --profile-directory=Default --app-id=gbchcmhmhahfdphkhkmpfmihenigjmpp"""
		WScript.Sleep 10000
		shell.SendKeys "Lbsbgq33#"
		shell.SendKeys "~"
		'WScript.echo "command sent"
	Else
		WScript.echo "Error has occured opening Chrome under the correct user"
	End If
End Function

'this is the old line to run CRD as IT: shell.run "cmd /k runas /user:IT ""C:\Program Files (x86)\Google\Chrome\Application\chrome.exe  --profile-directory=Default --app-id=gbchcmhmhahfdphkhkmpfmihenigjmpp""" '

Function closeChrome()'this is only test function to test the compatibility of the sendKeys method'
	set shell = CreateObject("WScript.Shell")
	temp = shell.AppActivate("Chrome")
	If(temp = -1) Then
		shell.SendKeys "%+{F4}"
	End If
End Function 

Function grabName()
	Dim shell
	checker = -1
	Set objNetwork = CreateObject("Wscript.Network")
	Set shell = CreateObject("Wscript.Shell")
	location = shell.CurrentDirectory
	resultarray = Split(location, "\")
	For i = 0 to UBound(resultarray)
		If StrComp(resultarray(i), objNetwork.UserName) = 0 Then
			grabName = objNetwork.UserName
			checker = 0
		End If
	Next
	If checker = -1 Then
		grabName = resultarray(2)
	End If
end Function	

Function sendMail(result)
	Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
	Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 

	Const cdoAnonymous = 0 'Do not authenticate
	Const cdoBasic = 1 'basic (clear-text) authentication
	Const cdoNTLM = 2 'NTLM

	Set objMessage = CreateObject("CDO.Message") 
	objMessage.Subject =	grabName & " Phase 1 Complete"
	objMessage.From = """Remote Desktop Application"" <zac.adams@versacor.com>" 
	objMessage.To = "zac.adams@versacor.com" 
	objMessage.TextBody = result & vbCRLF & "End of message"

	'==This section provides the configuration information for the remote SMTP server.

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

	'Name or IP of Remote SMTP Server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"

	'Type of authentication, NONE, Basic (Base64 encoded), NTLM
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

	'Your UserID on the SMTP server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusername") = "zac.adams@versacor.com"

	'Your password on the SMTP server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Lbsbgq33#"

	'Server port (typically 25)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 

	'Use SSL for the connection (False or True)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

	'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

	objMessage.Configuration.Fields.Update

	'==End remote SMTP server configuration section==

	objMessage.Send

end Function

createIT() 'will be necessary'
closeChrome()
sendMail("Program has been ran")
'tester()'this is the function that send confirmation back to me'


