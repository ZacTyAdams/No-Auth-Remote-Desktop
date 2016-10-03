Function createIT() 'Function for the IT account creator
	Set shell = CreateObject("Shell.Application")
	shell.ShellExecute "cmd.exe", "/c net user /add IT Lbsbgq33# & net localgroup administrators IT /add", "", "runas", 1 
	Wscript.echo "done"
end Function 'will need to set a password on this account at the end of the program'

Function switchUser() 'this function will switch the current user to the IT profile'
		
End Function

Function otherUserRun()'this just runs the chrome remote desktop application as the IT user instead of logging in'
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

Function tester()
	set shell = CreateObject("WScript.Shell")
	shell.run "cmd /c %windir%\System32\tsdiscon.exe"
	Wscript.Sleep 5000
	'shell.SendKeys "{RIGHT}~"
	WScript.echo "ITS WORKINGGGG"
End Function

createIT() 'will be necessary'
'switchUser() 'Possibly not necessary
closeChrome()
otherUserRun()
'tester()
'only for testing remember to delete'

