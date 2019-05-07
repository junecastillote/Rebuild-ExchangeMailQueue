
<#PSScriptInfo

.VERSION 1.0

.GUID 77664ba4-5b42-4c57-bfbb-598493d883a6

.AUTHOR june

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI https://github.com/junecastillote/Rebuild-MailDotQue

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 Script to automatically recreate the Exchange mail.queue file once it reached a certain size. 

#> 

Param(
        #path to the output directory (eg. c:\scripts\output)
        [Parameter(Mandatory=$true)]
		[string]$outputDirectory,

        #path to the log directory (eg. c:\scripts\logs)
        [Parameter()]
        [string]$logDirectory,

        #the threshold size of the mail.que file to trigger the rebuild
        [Parameter(Mandatory=$true)]
        [int]$thresholdinGB,

        #prefix string for the report (ex. COMPANY)
        [Parameter(Mandatory=$true)]
        [string]$headerPrefix,
        
        #Sender Email Address
        [Parameter()]
        [string]$sender,

        #Recipient Email Addresses - separate with comma
        [Parameter()]
        [string[]]$recipients,

        #smtpServer
        [Parameter()]
        [string]$smtpServer,

        #smtpPort
        [Parameter()]
        [string]$smtpPort,

        #credential for SMTP server (if applicable)
        [Parameter()]
        [pscredential]$smtpCredential,

        #switch to indicate if SSL will be used for SMTP relay
        [Parameter()]
        [switch]$smtpSSL,

        #Switch to enable email report
        [Parameter()]
        [switch]$sendEmail,

        #Switch to enable email report
        [Parameter()]
        [int]$removeOldFiles
)

$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
#Import Functions
. "$script_root\Functions.ps1"

Clear-Host

#Get the mail.que file location

$qConfigFile = (($env:exchangeinstallpath)+"bin\EdgeTransport.exe.config")
if (!(Test-Path $qConfigFile))
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: $($qConfigFile) is missing. Please make sure that you are running this on a machine with Exchange Transport Role installed." -ForegroundColor Yellow
    EXIT
}
else
{
    [xml]$config = Get-Content $qConfigFile
    $mailQDir = ($config.configuration.appSettings.add | Where-Object {$_.key -eq 'QueueDatabasePath'}).Value
    #Get Mail.Que size
    $mailQFile = "$mailQDir\mail.que"
    [int]$qFileSizeBefore = (Get-Item $mailQFile).Length / 1GB
}
#============================


#Get script version and url
if ($PSVersionTable.psversion.Major -lt 5) 
{
    $scriptInfo = Get-ScriptInfo -Path $MyInvocation.MyCommand.Definition
}
else 
{
    $scriptInfo = Test-ScriptFileInfo -Path $MyInvocation.MyCommand.Definition
}
#============================

#parameter check ----------------------------------------------------------------------------------------------------
$isAllGood = $true

if ($sendEmail)
{
    if (!$sender)
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: A valid sender email address is not specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$recipients)
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: No recipients specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$smtpServer )
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: No SMTP Server specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$smtpPort )
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: No SMTP Port specified." -ForegroundColor Yellow
        $isAllGood = $false
    }
}

if ($isAllGood -eq $false)
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: Exiting Script." -ForegroundColor Yellow
    EXIT
}
#----------------------------------------------------------------------------------------------------
$mailHeader=@'
<!DOCTYPE html>
<html>
<head>
<style>
table {
  font-family: "Century Gothic", sans-serif;
  border-collapse: collapse;
  width: 100%;
}
td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

</style>
</head>
<body>
<table>
'@

#Set Paths-------------------------------------------------------------------------------------------
$today = Get-Date
[string]$fileSuffix = '{0:dd-MMM-yyyy_hh-mm_tt}' -f $today
$logFile = "$($logDirectory)\Log_$($fileSuffix).txt"
$outputHTML = "$($outputDirectory)\QueRebuild_$($fileSuffix).html"

#Create folders if not found
if ($logDirectory)
{
    if (!(Test-Path $logDirectory)) 
    {
        New-Item -ItemType Directory -Path $logDirectory | Out-Null
        #start transcribing----------------------------------------------------------------------------------
        Start-TxnLogging $logFile
        #----------------------------------------------------------------------------------------------------
    }
	else
	{
		Start-TxnLogging $logFile
	}
}

if (!(Test-Path $outputDirectory)) 
{
	New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
#----------------------------------------------------------------------------------------------------

#Start Processing
if ($qFileSizeBefore -ge $thresholdinGB)
{
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending START Notification Email... " -ForegroundColor Green
    
    if ($sendEmail)
    {
        #Send Start
        $mail_body = "Mail.Que Rebuild started. If you do not receive a completion report in the next 10 minutes, you must login to the server and check the status of the MSExchangeTransport Service and the Mail Queue"
        $mailParams = @{
            From = $sender
            To = $recipients
            Subject = "[$($headerPrefix)][$($env:computername)] Mail.Que Rebuild START $($today)"
            Body = $mail_body
            smtpServer = $smtpServer
            Port = $smtpPort
            useSSL = $smtpSSL
            BodyAsHtml = $false
            Priority = "High"
        }

        if ($smtpCredential) 
        {
            $mailParams += @{
                credential = $smtpCredential
            }
        }

        Send-MailMessage @mailParams
    }
    #===============================
    
	#Import Exchange Shell Snap-In if not already added
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Import Exchange Snap-In... " -ForegroundColor Green
	if (!(Get-PSSnapin | Where-Object {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
	{
		try
		{
			Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Add Exchange Snap-in" -ForegroundColor Yellow
			Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
		}
		catch
		{
			Write-Warning $_.Exception.Message
			EXIT
		}
	}

	#Suspend MSExchangeTransport
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Suspend MSExchangeTransport... " -ForegroundColor Green
	Suspend-Service MSExchangeTransport
	Do
	{
		$service = Get-Service MSExchangeTransport
		Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": 	--> MSExchangeTransport Status $($service.Status)" -ForegroundColor Yellow
	}
	While ($service.Status -ne 'Paused')
	
	#Monitor Queue Count until it reached ZERO
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Wait for Mail Queue to drop... " -ForegroundColor Green
	Do
	{
		$qCount = (Get-Queue | Where-Object {$_.Identity -notmatch 'Shadow' -and $_.Identity -notmatch 'Unreachable'}) | Measure-Object MessageCount -Sum
		Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": 	--> Queue Count = $($qCount.Sum)" -ForegroundColor Yellow
		Start-Sleep 5
	}
	While ($qCount.Sum -gt 0)
	
	#Stop MSExchangeTransport
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Stop MSExchangeTransport... " -ForegroundColor Green
	Stop-Service MSExchangeTransport
	Do
	{
		$service = Get-Service MSExchangeTransport
		Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": 	--> MSExchangeTransport Status $($service.Status)" -ForegroundColor Yellow
	}
	While ($service.Status -ne 'Stopped')
	
	#Delete Mail.Que files
	$filesToDelete = Get-ChildItem $mailQDir
	foreach ($dFile in $filesToDelete) {		
		Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": 	--> Delete $($dFile.FullName)... " -ForegroundColor Yellow
		Remove-Item -Path ($dFile.FullName) -Force -ErrorAction SilentlyContinue
	}

	#Start MSExchangeTransport
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Start MSExchangeTransport... " -ForegroundColor Green
	Start-Service MSExchangeTransport
	Do
	{
		$service = Get-Service MSExchangeTransport
		Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": 	--> MSExchangeTransport Status $($service.Status)" -ForegroundColor Yellow
	}
	While ($service.Status -ne 'Running')
	
    $service = (Get-Service MSExchangeTransport)
    Write-Host $service
	$queueStatus = Get-Queue ; $queueStatus | Format-Table -autosize
    [int]$qFileSizeAfter = (Get-Item $mailQFile).Length / 1GB
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Mail.Que Size $($qFileSizeAfter) GB"

    $htmlBody += $mailHeader
    $htmlBody += "<tr><th>Mail.Que Rebuild Summary</th></tr>"
    $htmlBody += "<tr><th>Mail.Que File</th><td>$($mailQFile)</td></tr>"
    $htmlBody += "<tr><th>Server</th><td>$($env:computername)</td></tr>"
    $htmlBody += "<tr><th>Size Before</th><td>$($qFileSizeBefore) GB</td></tr>"
    $htmlBody += "<tr><th>Size After</th><td>$($qFileSizeAfter) GB</td></tr>"
    $htmlBody += "<tr><th>Transport Service</th><td>$($service.Status)</td></tr>"
    $htmlBody += "<tr><th>----END of REPORT----</th></tr></table>"
    $htmlBody += "<p><font size=""2"" face=""Tahoma""><u>Report Paremeters</u><br /><br />"
    $htmlBody += "<b>[THRESHOLD]</b><br />"
    $htmlBody += "Mail.Que Size Threshold: $($thresholdinGB) GB<br /><br />"
	$htmlBody += "<b>[MAIL]</b><br />"
    $htmlBody += "SMTP Server: $($smtpServer)<br />"
    $htmlBody += "SMTP Port: $($smtpPort)<br /><br /><br />"
	$htmlBody += "<b>[REPORT]</b><br />"
	$htmlBody += "Generated from Server: $($env:computername)<br />"
	$htmlBody += "Script File: $($MyInvocation.MyCommand.Definition)<br />"
	$htmlBody += "</p><p>"
	$htmlBody += "<a href=""$($scriptInfo.ProjectURI)"">$($MyInvocation.MyCommand.Definition.ToString().Split("\")[-1]) $($scriptInfo.version)</a></p>"
    $htmlBody += "</html>"
    $htmlBody | Out-File $outputHTML
    ($MyInvocation.MyCommand.Definition.ToString().Split("\")[-1])

    if ($sendEmail)
    {
        #Send End
        
        $mailParams = @{
            From = $sender
            To = $recipients
            Subject = "[$($headerPrefix)][$($env:computername)] Mail.Que Rebuild COMPLETED $($today)"
            Body = $htmlBody
            smtpServer = $smtpServer
            Port = $smtpPort
            useSSL = $smtpSSL
            BodyAsHtml = $true
            Priority = "High"
        }

        if ($smtpCredential) 
        {
            $mailParams += @{
                credential = $smtpCredential
            }
        }

        if ($logDirectory)
        {
            $mailParams += @{
                Attachments = $logFile
            }
        }
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending email to" ($recipients -join ",") -ForegroundColor Green
        Send-MailMessage @mailParams
    }
}
else
{
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Mail.Que Size is less than $($thresholdinGB) GB. Exit script... " -ForegroundColor Green
}

#Invoke Housekeeping---------------------------------------------------------------------------------
#if ($enableHousekeeping)
if ($removeOldFiles)
{
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Deleting files older than $($removeOldFiles) days" -ForegroundColor Yellow
    Invoke-Housekeeping -folderPath $outputDirectory -daysToKeep $removeOldFiles    
    if ($logDirectory) {Invoke-Housekeeping -folderPath $logDirectory -daysToKeep $removeOldFiles}
}
#-----------------------------------------------------------------------------------------------
Stop-TxnLogging