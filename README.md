# Get-EnabledProtocolReport.ps1

Get a list of mailbox users having a selected client access protocol enabled

## Description

This scripts gather a list of enabled users for a selected Exchange Server client protocol. The list of users is sent by email as HTML text in the email body or as an attached CSV file. You can select to gather data for a single protocol or for all protocols.

Available protocols are:

- POP
- IMAP
- ActiveSync

## Requirements

- Windows Server 2012 R2 or newer
- Exchange 2016+ Management Shell
- GlobalFunctions module ([found here](http://scripts.granikos.eu))

## Parameters

### Protocol

The client access protocol to report on.
Options: All, POP, IMAP, ActiveSync

### ExportCsv

Switch to export the result set of users as CSV file and attach the file to email

### UserDetailsInEmailBody

Switch to include the result set of users in the mail body

### SendMail

Switch to send a report email

### MailFrom

Email address of report sender

### MailTo

Email address of report recipient

### MailServer

SMTP Server for email report

## Examples

``` PowerShell
.\Get-EnabledProtocolReport.ps1 -SendMail -MailFrom automation@varunagroup.de -MailTo report@varunagroup.de -MailServer relay.varunagroup.de -Protocol ALL
```

Find users having all protocols enabled, create a CSV file per protocol and send an email with CSV attachments

``` PowerShell 
.\Get-EnabledProtocolReport.ps1 -Protocol ALL -ExportCsv
```

Find users having all protocols enabled, create a CSV file per protocol

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery

Download and vote at TechNet Gallery

- [https://gallery.technet.microsoft.com/Get-a-list-of-mailbox-29cd5ba3](https://gallery.technet.microsoft.com/Get-a-list-of-mailbox-29cd5ba3)

## Credits

Written by: Thomas Stensitzki

Stay connected:

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)
