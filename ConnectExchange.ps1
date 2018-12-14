function Connect-Exchange
{
    begin
    {
        $ShouldConnect = $false
        $ExecutionPolicy = Get-ExecutionPolicy
        if ($ExecutionPolicy -ne "RemoteSigned") {
            Write-Warning "To run this script, the Execution Policy must be set to RemoteSigned. Currently is it $ExecutionPolicy. Would you like to change it?"

            $Confirm = Read-Host "Type Y or N"
            if ($Confirm -eq "Y") {
                Start-Process powershell -Verb runAs "Set-ExecutionPolicy RemoteSigned"
                $ShouldConnect = $true
            }
        } else {
            $ShouldConnect = $true
        }
        if ($ShouldConnect) {
            Write-Warning "Connecting to Exchange."
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "EMAIL SERVER" -Authentication Kerberos
            Import-Module (Import-PSSession $Session -AllowClobber) -Global
            Write-Warning "Connected to Exchange. Use 'Get-PSSession | Remove-PSSession' to clean up your Exchange session."
            $idFromSession = $Session.id
        }
        return $idFromSession
    }
}