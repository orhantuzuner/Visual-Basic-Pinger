Option Explicit

Dim Targets, Target

Targets=Array(_
	"google.com",_
	"192.168.1.2:8080"_
	)
	
For Each Target In Targets
  
  Ping Target
  
Next


' pinger dunction
Function Ping( TargetHost )
	
	Dim Shell, StrCommand, ReturnCode, Parameters, CustomText
	Set Shell = CreateObject("wscript.shell")
	
	' if target host contain a port
	If InStr(TargetHost,":") > 0 Then
		TargetHost = Replace(TargetHost,":"," ")
		
		' shell command string
		StrCommand = "tcping.exe -n 1 " & TargetHost
	Else' if not contain a port
		' shell command string
		StrCommand = "ping -n 1 -w 300 " & TargetHost
	End if
	
	' return val 0 = true | else = false
	ReturnCode = Shell.Run(StrCommand,0,TRUE)
	
	
	' if ping is correct
	If ReturnCode = 0 Then
		' do something
	End If
	
	Set Shell = Nothing
End Function

