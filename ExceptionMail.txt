
function SendMail(args)


On error resume Next

Dim objMessage, objConfig, fields
Dim filePath,UsrName,FromEmail,ToEmail,Password,smtpServer,smtpPort


filePath = args(0)
UsrName = args(1)
Password  = args(2)
FromEmail = args(3)
ToEmail = args(4)
smtpServer =args(5)
smtpPort = args(6)
mailBody = args(7)
mailSubject = args(8)

  msgBox filePath & "\n" & mailBody & "\n" & mailSubject 
    'Create message 
    Set objMessage = CreateObject("CDO.Message")
    Set objConfig = CreateObject("CDO.Configuration")
    Set fields = objConfig.Fields

    With fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") =smtpPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = UsrName
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Password 'Enter Your Gmail APP Password
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Update
    End With

    Set objMessage.Configuration = objConfig

    objMessage.From = FromEmail
    objMessage.To = ToEmail
    objMessage.Subject = mailSubject
    objMessage.TextBody = mailBody
  
	If Trim(filePath) <> "" Then
   		objMessage.AddAttachment filePath
   	End If
    
    objMessage.Send



If Err.Number <> 0 Then
	SendMail = "Error Number is: " & Err.Number & "Error Description is: " & Err.Description
	msgBox SendMail
	Err.Clear
Else
	SendMail = "Success"
	msgBox SendMail
	
End If


On Error GoTo 0

End function