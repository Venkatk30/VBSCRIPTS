'Option Explicit

function EmailGenerator(args)


On error resume Next

Dim objMessage, objConfig, fields
Dim folderPath, fso, folder, file
Dim subjects, bodies
Dim subjectLine, bodyText
Dim randIndex
Dim UsrName,FromEmail,ToEmail,Password,smtpServer,smtpPort

'folderPath = "C:\Users\venkat\Documents\Automation Anywhere\Bot Development\My Docs\Input"   'Folder containing statements
folderPath = args(0)
UsrName = args(1)
Password  = args(2)
FromEmail = args(3)
ToEmail = args(4)
smtpServer =args(5)
smtpPort = args(6)
'Random subjects
subjects = Array( _
"Reconciliation Statement Ref#", _
"Bank Statement Submission ID#", _
"Account Statement Processing#", _
"Daily Reconciliation File#", _
"Finance Statement Record#" _
)

'Random mail body
bodies = Array( _
"Please process the attached bank statement for reconciliation.", _
"Attached file contains the account statement for verification.", _
"Kindly review the enclosed document for reconciliation.", _
"This file is submitted for financial reconciliation processing.", _
"Attached statement requires validation in reconciliation system." _
)

Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(folderPath)

Randomize

For Each file In folder.Files

    'Create message each loop
    Set objMessage = CreateObject("CDO.Message")
    Set objConfig = CreateObject("CDO.Configuration")
    Set fields = objConfig.Fields

    With fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer'"smtp.gmail.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") =smtpPort' 465
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = UsrName'"venkat.kusuma@cxdatalabs.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Password'"avjq wndj kvxo zbqj"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Update
    End With

    Set objMessage.Configuration = objConfig

    'Generate random subject and body
    randIndex = Int((UBound(subjects)+1) * Rnd)
    subjectLine = subjects(randIndex) & " " & Int(100000 * Rnd)

    randIndex = Int((UBound(bodies)+1) * Rnd)
    bodyText = bodies(randIndex)

    objMessage.From = FromEmail'"venkat.kusuma@cxdatalabs.com"
    objMessage.To = ToEmail'"venkat.kusuma@cxdatalabs.com"
    objMessage.Subject = subjectLine
    objMessage.TextBody = bodyText

    objMessage.AddAttachment file.Path

    objMessage.Send

Next


If Err.Number <> 0 Then
	EmailGenerator = "Error Number is: " & Err.Number & "Error Description is: " & Err.Description
	msgBox EmailGenerator
	Err.Clear
Else
	EmailGenerator = "Success"
	msgBox EmailGenerator
	
End If


On Error GoTo 0

End function