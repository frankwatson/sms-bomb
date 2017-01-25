'Created by Frank Watson
'github.com/frankwatson

do
on error resume next

Const schema   = "http://schemas.microsoft.com/cdo/configuration/"
Const cdoBasic = 1
Const cdoSendUsingPort = 2
Dim oMsg, oConf

' E-mail properties
Set oMsg      = CreateObject("CDO.Message")
oMsg.From     = "spam@gmail.com"
oMsg.To       = "CHANGEME" 'Refer to http://bit.ly/2kj7CwW
oMsg.Subject  = "CHANGEME"
oMsg.TextBody = "CHANGEME"

' GMail SMTP server configuration and authentication info
Set oConf = oMsg.Configuration
oConf.Fields(schema & "smtpserver")       = "smtp.gmail.com" 'server address
oConf.Fields(schema & "smtpserverport")   = 465              'port number
oConf.Fields(schema & "sendusing")        = cdoSendUsingPort
oConf.Fields(schema & "smtpauthenticate") = cdoBasic         'authentication type
oConf.Fields(schema & "smtpusessl")       = True             'use SSL encryption
oConf.Fields(schema & "sendusername")     = "CHANGEME" 'sender username (CHANGE THIS)
oConf.Fields(schema & "sendpassword")     = "CHANGEME"      'sender password (CHANGE THIS)
oConf.Fields.Update()

' send message
oMsg.Send()
loop

' Return status message
If Err Then
	resultMessage = "ERROR " & Err.Number & ": " & Err.Description
	Err.Clear()
Else
	resultMessage = "Message sent ok"
End If

Wscript.echo(resultMessage)
