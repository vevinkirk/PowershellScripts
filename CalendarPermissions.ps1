function Set-CalendarPermissions{

        <#
        .SYNOPSIS
            Adds calendar permissions to a calendar for a given user

        .DESCRIPTION
            Grants permissions for a user for a calendar

        .PARAMETER employeeFirstLast
            The user to add permissions for format(first.last)

        .PARAMETER calendarEmail
            The calendar to add user permissions for format(calendarEmail)

        .PARAMETER permissionLevel
            The user to add permissions for format(first.last)

        .EXAMPLE
            PS C:\> Set-CalendarPermissions -Employee jane.doe -Calendar fred.smith -Permission Owner
        #>

        Param(
            [Parameter(Mandatory=$true)]
            [string]$Employee,
            [Parameter(Mandatory=$true)]
            [string]$Calendar,
            [Parameter(Mandatory=$true)]
            [string]$PermissionLevel
            ) #end param

        if((Get-PSSession -ErrorAction silentlycontinue) -ne $null){ #check if connected to exchange
          Write-Verbose "Connected to Exchange!"
        }
        else{
          $idFromSession = Connect-Exchange #connect if not connected
        }

        Write-Warning "Adding Permissions"

        if((Get-Mailbox $Employee -ErrorAction silentlycontinue) -ne $null){ #check if mailbox exists
             Write-Verbose "Mailbox exists"
        }
        else
        {
           Throw "Could not find employee matching $Employee " 
        }

        if((get-mailbox $Calendar -ErrorAction silentlycontinue) -ne $null){
            Write-Verbose "Maibox exists"
        }
        else
        {
           Throw "Could not find matching calendar for $Calendar "
        }

        $Calendar = $Calendar + ":\calendar"

        Add-MailboxFolderPermission $Calendar -user $Employee -AccessRights $PermissionLevel

        if($idFromSession -ne $null){
            Remove-PSSession $idFromSession #cleaning up session
        }
}



function Remove-CalendarPermissions{
      <#
      .SYNOPSIS
          Remove calendar permissions to a calendar for a given user

      .DESCRIPTION
          Remove permissions for a user for a calendar

      .PARAMETER employeeFirstLast
          The user to remove permissions for format(first.last)

      .PARAMETER calendarEmail
          The calendar to remove user permissions for format(calendarEmail)

      .EXAMPLE
          PS C:\> Remove-CalendarPermissions -Employee jane.doe -Calendar fred.smith
      #>
        Param(
            [Parameter(Mandatory=$true)]
            [string]$Employee,
            [Parameter(Mandatory=$true)]
            [string]$Calendar
            ) #end param

        if((Get-PSSession -ErrorAction silentlycontinue)-ne $null){
          write-verbose "Connected to Exchange!"
        }
        else{
          $idFromSession = connect-exchange
        }

        Write-Warning "Please use Connect-Exchange CMDlet first in order to be connected to exchange"
        Write-Warning "Removing Permissions"

        if((Get-Mailbox $Employee -ErrorAction silentlycontinue) -ne $null){
             Write-Verbose "Mailbox exists"
        }
        else
        {
           Throw "Could not find employee matching $Employee "
        }

        if((get-mailbox $Calendar -ErrorAction silentlycontinue) -ne $null){
            Write-Verbose "Maibox exists"
        }
        else
        {
           Throw "Could not find matching calendar for $Calendar "
        }

        $Calendar = $Calendar + ":\calendar"

        Remove-MailboxFolderPermission $Calendar -user $Employee

        if($idFromSession -ne $null){
            Remove-PSSession $idFromSession #cleaning up session
        }
}


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
