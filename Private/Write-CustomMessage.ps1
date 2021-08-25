using namespace System.Management.Automation
#Requires -Version 7
function Write-CustomMessage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [string]$Message,
        [Parameter(Mandatory = $True)]
        [string]$Type
    )
    Switch ($Type) {
        'Success' {
            $ForegroundColour = 'Green'
            $Prefix = 'SUCCESS: '
        }
        'Information' {
            $ForegroundColour = 'Blue'
            $Prefix = 'INFO: '
        }
    }
    $MessageData = [HostInformationMessage]@{
        Message = "$($Prefix)$($Message)"
        ForegroundColor = $ForegroundColour
    }
    Write-Information -MessageData $MessageData -InformationAction Continue
}