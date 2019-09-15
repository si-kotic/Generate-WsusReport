Param (
    [CmdletBinding()]
    [string]$customerName,
    [Parameter(Mandatory)][string][ValidateScript(
        IF ($_ -match '(^[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+$)') {
            $true
        } ELSE {
            Throw "$_ is not a valid email address"
        }
    )]$wsusServer,
    [Parameter(Mandatory)][string][ValidateScript(
        IF ($_ -match '(^[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+$)') {
            $true
        } ELSE {
            Throw "$_ is not a valid email address"
        }
    )]$dc,
    [string]$retiredMachineOU = "Retired",
    [Parameter(Mandatory)][string][ValidateScript(
        IF ($_ -match '(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)') {
            $true
        } ELSE {
            Throw "$_ is not a valid email address"
        }
    )]$FromAddress,
    [Parameter(Mandatory)][string][ValidateScript(
        IF ($_ -match '(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)') {
            $true
        } ELSE {
            Throw "$_ is not a valid email address"
        }
    )]$ToAddress,
    [Parameter(Mandatory)][string][ValidateScript(
        IF ($_ -match '(^[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+$)') {
            $true
        } ELSE {
            Throw "$_ is not a valid email address"
        }
    )]$SmtpServer = "smtp.office365.com",
    [ValidateSet(25,587,465,2525)]$SmtpPort = 587,
    [string]$EmailSubject = "WSUS Report"
)

$wsusSession = New-PSSession $wsusServer
Invoke-Command -Session $wsusSession -ScriptBlock {
    Import-Module UpdateServices
}
$wsusComputers = Invoke-Command -Session $wsusSession -ScriptBlock {
    Get-WsusComputer | Where-Object {$_.LastSyncTime -lt (Get-Date).AddMonths(-1)}
}
$newUpdates = Invoke-Command -Session $wsusSession -ScriptBlock {
    Get-WsusUpdate | Where-Object {$_.Approved -eq "NotApproved" -and $_.UpdatesSupersedingThisUpdate[0] -eq "None"} | Select-Object @{N="Update Name";E={$_.Update.Title}},Products,Classification
}
$wsusProducts = Invoke-Command -Session $wsusSession -ScriptBlock {
    Get-WsusProduct | Where-Object {$_.Product.ArrivalDate -gt (Get-Date).AddMonths(-1)}
}

$dcSession = New-PSSession $dc
Invoke-Command -Session $dcSession -ScriptBlock {
    Import-Module ActiveDirectory
}
$wsusReport = $wsusComputers | Foreach-Object {
    $curObj = $_
    Write-Debug -Message "Current Computer from WSUS is $($curObj.FullDomainName)"
    Invoke-Command -Session $dcSession -ArgumentList $curObj,$retiredMachineOU -ScriptBlock {
        Param (
            $curObj,
            $retiredMachineOU
        )
        $Report = "" | Select-Object HostName,LastContactDate,RecommendedAction
        $curComputer = Get-ADComputer $curObj.FullDomainName.Split(".")[0]
        Write-Debug -Message "Current Computer from AD is $($curComputer.DNSHostName)"
        Write-Debug -Message "Computer Distinguished Name = $($curComputer.DistinguishedName)"
        Write-Debug -Message "Computer AD Account Enabled:  $($curComputer.Enabled)"
        IF ($curComputer.DistinguishedName -like "*$retiredMachineOU*" -or !$curComputer.Enabled) {
            Write-Debug -Message "Computer identified as Retired"
            $Report.HostName = $curObj.FullDomainName
            $Report.LastContactDate = $curObj.LastSyncTime
            $Report.RecommendedAction = "Machine retired.  Delete from WSUS."
            $Report
        } ELSE {
            Write-Debug -Message "Computer identified as Active"
            $Report.HostName = $curObj.FullDomainName
            $Report.LastContactDate = $curObj.LastSyncTime
            $Report.RecommendedAction = "Turn on machine and check for updates."
            $Report
        }
    }
} | Select-Object HostName,LastContactDate,RecommendedAction
Remove-PSSession -Session $wsusSession
Remove-PSSession -Session $dcSession

IF ($wsusReport) {
    [string]$computersHTML = $wsusReport | ConvertTo-Html -Fragment
} ELSE {
    [string]$computersHTML = "All machines have checked in within the last month."
}
IF ($newUpdates) {
    [string]$updatesHTML = $newUpdates | Select-Object "Update Name",Products,Classification | ConvertTo-Html -Fragment
} ELSE {
    [string]$updatesHTML = "There are no unapproved updates to report."
}
IF ($wsusProducts) {
    [string]$productsHTML = $wsusProducts | ConvertTo-Html -Fragment
} ELSE {
    [string]$productsHTML = "No new products have been released in the last month."
}

$mailBody = @"
<html>
<head>
<style>
body {
font-family: verdana
}
h1 {
background-color: deepskyblue;
padding: 7px;
}
table {
border: 1px solid #ddd;
border-collapse: collapse;
}
th {
padding-top: 12px;
padding-bottom: 12px;
text-align: left;
border: 1px solid #ddd;
background-color: deepskyblue;
padding: 7px;
}
td {
border: 1px solid #ddd;
padding: 7px;
}
tr:nth-child(even) {
background-color: #f2f2f2;
}
</style>
</head>
<body>
<h1>Updates released in the last month:</h1>
$productsHTML
<h1>Machines which have not checked in during the last month:</h1>
$computersHTML
<h1>Updates which are not approved:</h1>
$updatesHTML
</body>
</html>
"@

Send-MailMessage -From $FromAddress -To $ToAddress -SMTPServer $SmtpServer -Subject $EmailSubject -BodyAsHTML $mailBody