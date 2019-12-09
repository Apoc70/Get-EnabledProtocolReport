<#
  .SYNOPSIS
  Get a list of mailbox users having a selected client access protocol enabled
   
  Thomas Stensitzki
	
  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
  Version 1.0, 2019-12-08

  Ideas, comments and suggestions to support@granikos.eu 
 
  .LINK  
  http://scripts.Granikos.eu
	
  .DESCRIPTION

  This script gathers a list of enabled users for a selected Exchange Server client protocol.
  The list of users is sent by email as HTML text in the email body or as an attached CSV file.
  You can select to gather data for a single protocol or for all protocols. 
	
  .NOTES 

  Requirements 
  - Windows Server 2012 R2 or newer
  - Exchange 2016+ Management Shell
  - GlobalFunctions module ((http://scripts.granikos.eu)

  Revision History 
  -------------------------------------------------------------------------------- 
  1.0     Initial community release 
	
  .PARAMETER Protocol
  The client access protocol to report on. 
  Options: All, POP, IMAP, ActiveSync

  .PARAMETER ExportCsv
  Switch to export the result set of users as CSV file and attach the file to email
  
  .PARAMETER UserDetailsInEmailBody
  Switch to include the result set of users in the mail body

  .PARAMETER SendMail
  Switch to send a report email

  .PARAMETER MailFrom
  Email address of report sender

  .PARAMETER MailTo
  Email address of report recipient

  .PARAMETER MailServer
  SMTP Server for email report
   
  .EXAMPLE
  .\Get-EnabledProtocolReport.ps1 -SendMail -MailFrom automation@varunagroup.de -MailTo report@varunagroup.de -MailServer relay.varunagroup.de -Protocol ALL -ExportCsv

  Find users having all protocols enabled, create a CSV file per protocol and send an email with CSV attachments

  .EXAMPLE
  .\Get-EnabledProtocolReport.ps1 -Protocol ALL -ExportCsv

  Find users having all protocols enabled, create a CSV file per protocol only
#>

[CmdletBinding()]
Param(
  [ValidateSet('All','POP','IMAP','ActiveSync')] 
  [string]$Protocol = 'All',
  [switch]$ExportCsv,
  [switch]$UserDetailsInEmailBody,
  [switch]$SendMail,
  [string]$MailFrom = '',
  [string]$MailTo = '',
  [string]$MailServer = ''
)

Add-Type -AssemblyName System.Web
Import-Module -Name GlobalFunctions
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$FileTimeStamp = $(Get-Date -Format 'yyyy-MM-dd HHmm')
$global:Files =@()
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 30
$logger.Purge()
$logger.Write('Script started')

Function Test-SendMail {
  if( ($MailFrom -ne '') -and ($MailTo -ne '') -and ($MailServer -ne '') ) {
    return $true
  }
  else {
    return $false
  }
}

function Get-CASMailboxForProtocol {
  [CmdletBinding()]
  param(
    [string]$Protocol
  )

  Write-Verbose -Message ('Fetching user mailboxes for {0}' -f $Protocol)

  $htmlResult = ('<h3>{0}</h3>' -f $Protocol)

  switch($Protocol) {
    'POP' {
      $result = Get-CASMailbox -ResultSize Unlimited | Where-Object{$_.PopEnabled -eq $true} | Select-Object -Property DisplayName,PrimarySmtpAddress | Sort-Object -Property DisplayName
    }
    'IMAP' {
      $result = Get-CASMailbox -ResultSize Unlimited | Where-Object{$_.ImapEnabled -eq $true} | Select-Object -Property DisplayName,PrimarySmtpAddress | Sort-Object -Property DisplayName
    }
    'ActiveSync' {
      $result = Get-CASMailbox -ResultSize Unlimited | Where-Object{$_.ActiveSyncEnabled -eq $true} | Select-Object -Property DisplayName,PrimarySmtpAddress | Sort-Object -Property DisplayName
    }
  }

  try{
    $count = ($result | Measure-Object).Count
    $text = ('Found {0} mailboxes for {1} protocol.' -f $count, $Protocol)
    $logger.Write($text)
    Write-Verbose -Message $text

    $htmlResult += ('<p>{0}</p>' -f ('Found <strong>{0}</strong> mailboxes for <strong>{1}</strong> protocol.' -f $count, $Protocol)) 

    if($ExportCsv) {
      $fileName = (Join-Path -Path $ScriptDir -ChildPath ('{0}-{1}.csv' -f $Protocol, $FileTimeStamp))
      $result | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Force -Confirm:$false
      if(Test-Path -Path $fileName){
        Write-Verbose -Message ('File exists: {0}' -f $fileName)
        $global:Files += $fileName
      }
    }

    if($UserDetailsInEmailBody) { 
      
      $htmlResult += '<table>'

      foreach ($object in $result) {
        $htmlResult += ('<tr><td>{0}</td><td>{1}</td></tr>' -f $object.DisplayName, $object.PrimarySmtpAddress)
      }
  
      $htmlResult += '</table><hr />'

    }

  }
  catch {}
  
  $htmlResult
}

# Main -----------------------------------------------------

If ($SendMail.IsPresent) { 
  If (-Not (Test-SendMail)) {
    Throw 'If -SendMail specified, -MailFrom, -MailTo and -MailServer must be specified as well!'
  }
}

### MAIN ###################################

# View entire Active Directory forest
Set-ADServerSettings -ViewEntireForest $true

$message = ('Fetching CAS mailboxes for {0} protocol(s).' -f ($Protocol))
$logger.Write($message)
Write-Verbose -Message $message

# Prepare Output
$Output = '<html>
  <body>
<font size=""1"" face=""Arial,sans-serif"">'

switch($Protocol) {
  'POP' {
    $Output += Get-CASMailboxForProtocol -Protocol 'POP'
  }

  'IMAP' {
    $Output += Get-CASMailboxForProtocol -Protocol 'IMAP'   
  }
  'ActiveSync' {
    $Output += Get-CASMailboxForProtocol -Protocol 'ActiveSync'   
  }
  'ALL' {
    $Output += Get-CASMailboxForProtocol -Protocol 'POP'
    $Output += Get-CASMailboxForProtocol -Protocol 'IMAP'
    $Output += Get-CASMailboxForProtocol -Protocol 'ActiveSync'   
  }
}

$Output += '</font></body></html>'

$Body = [Web.HttpUtility]::HtmlDecode($Output)

if($SendMail) {
  
  try {
    if($ExportCsv) {
      $logger.Write(('Sending email with attachments to {0}' -f $MailTo))
      Send-MailMessage -Encoding utf8 -From $MailFrom -To $MailTo -Subject ('Get-EnabledProtocolReport - {0}' -f $Protocol) -SmtpServer $MailServer -BodyAsHtml -Body $Body -Attachments $global:Files
    }
    else { 
      $logger.Write(('Sending email to {0}' -f $MailTo))
      Send-MailMessage -Encoding utf8 -From $MailFrom -To $MailTo -Subject ('Get-EnabledProtocolReport - {0}' -f $Protocol) -SmtpServer $MailServer -BodyAsHtml -Body $Body 
    }
  }
  catch {
    $logger.Write(('Error sending email to {0}' -f $MailTo),3)
  }
}

$message = 'Script finished.'
$logger.Write($message)
Write-Verbose -Message $message