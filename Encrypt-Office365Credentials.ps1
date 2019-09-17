Param (
    [Parameter(DontShow)]$Office365Credentials = (Get-Credential),
    [Parameter(Mandatory)]$Path
)
$Office365Credentials | Export-CliXml -Path $Path