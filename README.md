# Generate-WsusReport
Generate-WsusReport will generate and email a report on the following:
* Updates released in the last month.
* Machines which have not checked in with WSUS in the last month.
* Updates which have neither been approved nor declined (and have not been superseded by a subsequent update).

## Scripts
* Generate-WsusReport.ps1
* Encrypt-Office365Credentials.ps1

## Usage
### Generate-WsusReport
Generate-WsusReport must be run as an Administrator as a user with access to both WSUS and Active Directory.  It connects remotely to the WSUS Server and the Domain Controller so Remote Management must be configured and enabled on these machines.
#### Parameters
##### CustomerName
This parameter is currently not used but is intended to identify the customer for whom the script is being run.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | False
##### WsusServer
Specify the FQDN of the WSUS server.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | True
##### DC
Specify the FQDN of the Domain Controller.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | True
##### RetiredMachineOU
Use this optional parameter to specify the name of an Organisational Unit in which retired machines are stored in Active Directory.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | Retired
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | False
##### FromAddress
Specify the email address from which the report is emailed.  If you are using Office 365 to send your email then this address must be the address used to authenticate with Office 365.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | True
##### ToAddress
Specify the email address to which the report is emailed.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | True
##### SmtpServer
Specify the SMTP Server through which to send the email.  It defaults to using Office 365.  Currently authentication is only supported for Office 365 but this can be addressed if it is required.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | smtp.office365.com
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | False
##### SmtpPort
Specify the port used to contact the SMTP Server.  Accepts the values 25,587,465,2525.

Argument | Value
--- | ---
Type | Integer
Position | Named
Default value | 587
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | False
##### MailSubject
Specify the subject of the email report to be generated.  It has a default value in the format `WSUS Report - Month Year - WsusServerName`

Argument | Value
--- | ---
Type | String
Position | Named
Default value | "WSUS Report - " + (Get-Date -Format "MMM yyyy") + " - $wsusServer"
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | False
##### Office365CredentialFile
Specify the location of a file containing Encrypted Office 365 Credentials.  You can create this file using the accompanying script `Encrypt-Office365Credentials.ps1`.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | False
#### Syntax
```powershell
Generate-WsusReport.ps1 -customerName "Customer Name" -wsusServer "wsusserver.domain.local" -dc "dc.domain.local" -FromAddress "sendingaddress@domain.com" -ToAddress "recipientaddress@anotherdomain.com" -Office365CredentialFile "C:\Path\To\Credentials\File.xml
```

### Encrypt-Office365Credentials
Encrypt-Office365Credentials can be used to generate an XML file containing encrypted credentials for use when using this script non-interactively and authenticating with an Office 365 Server.  The XML file can only be decrypted by the user who created it, on the machine on which it was created.
#### Paramters
##### Office365Credentials
**This parameter should not be explicitly specified when calling the cmdlet and will not autocomplete.  When you execute the command without this paramter you will be prompted to provide it and the value will be masked for security.**

Argument | Value
--- | ---
Type | PSCredential
Position | Hidden
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | False
##### Path
Specify the filepath to output the encrypted credentials.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | True
#### Syntax
```powershell
Encrypt-Office365Credentials.ps1 -Path "C:\Path\To\Credentials\File.xml"
```