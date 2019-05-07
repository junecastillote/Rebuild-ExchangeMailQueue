#EDIT THESE VALUES

#[MANDATORY]
#Where is your script file (.PS1) located?
$scriptFile = "E:\Scripts\Rebuild-MailDotQue\Rebuild-MailDotQue.PS1"

#[MANDATORY]
#Where do we save the output?
$outputDirectory = "E:\Scripts\Rebuild-MailDotQue\Output"

#[MANDATORY]
#String you want to show in the Title or Subject (eg. Company Name)
$headerPrefix = "COMPANY"

#[MANDATORY]
#Threshold in GB for the Mail.Que filesize
$thresholdinGB = 10

#[OPTIONAL]
#Where do we put the transcript log?
$logDirectory = "E:\Scripts\Rebuild-MailDotQue\Log"

#======================================
#start EMAIL SECTION
#======================================
#[OPTIONAL]
#Do you want to send the email summary? $true or $false
$sendEmail = $true

#[REQUIRED IF $sendEmail = $true]
#If we will send the email summary, what is the sender email address we should use?
#This must be a valid, existing mailbox and address in Office 365
#The account you use for the Credential File must have "Send As" permission on this mailbox
$sender = "Sender Name <sender@domain.com>"

#[REQUIRED IF $sendEmail = $true]
#Who are the recipients?
#Multiple recipients can be added (eg. "recipient1@domain.com","recipient2@domain.com")
$recipients = "recipient1@domain.com","recipient2@domain.com"

#[REQUIRED IF $sendEmail = $true]
#your SMTP relay server
$smtpServer = "smtp.server.here"

#[REQUIRED IF $sendEmail = $true]
#your SMTP relay server port
$smtpPort = "25"

#[OPTIONAL - use only if your SMTP Relay requires authentication]
#Which XML file contains your SMTP relay authentication? - IF APPLICABLE
#If you don't have this yet, run this: Get-Credential | Export-CliXML <file.xml>
#Or if you are using the same account to login to Office 365, just point to the same XML file
$smtpCredentialFile = ""

#[OPTIONAL - use only if SMTP Relay requires SSL]
#Indicate whether or not SSL will be used
$smtpSSL = $false
#======================================
#end EMAIL SECTION
#======================================

#[OPTIONAL]
#If you want to delete older backups, define the age in days.
$removeOldFiles = 30

#======================================
#DO NOT TOUCH THE BELOW CODES
#======================================
if ($smtpCredentialFile) {$smtpCredential = Import-Clixml $smtpCredentialFile}

$params = @{
    outputDirectory = $outputDirectory    
    sendEmail = $sendEmail
    smtpSSL = $smtpSSL
	thresholdinGB = $thresholdinGB
	headerPrefix = $headerPrefix
}

if ($removeOldFiles) {$params += @{removeOldFiles = $removeOldFiles}}
if ($logDirectory) {$params += @{logDirectory = $logDirectory}}
if ($smtpServer) {$params += @{smtpServer = $smtpServer}}
if ($sender) {$params += @{sender = $sender}}
if ($recipients) {$params += @{recipients = $recipients}}
if ($smtpPort) {$params += @{smtpPort = $smtpPort}}
if ($smtpCredentialFile)  {$params += += @{smtpCredential = $smtpCredential}}
#======================================

& "$scriptFile" @params