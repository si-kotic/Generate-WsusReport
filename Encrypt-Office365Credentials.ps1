Param (
    [Parameter(DontShow)]$Office365Credentials = (Get-Credential),
    $Path
)
$Office365Credentials | Export-CliXml -Path $Path