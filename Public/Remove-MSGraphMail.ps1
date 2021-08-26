function Remove-MSGraphMail {
    [CmdletBinding( SupportsShouldProcess = $True, ConfirmImpact = 'High' )]
    param (
        # Specify the mailbox (or UPN) to remove an email for.
        [Parameter(Mandatory = $true, ParameterSetName = 'Single', ValueFromPipelineByPropertyName)]
        [string]$Mailbox,
        # The ID of the message to remove.
        [Parameter(Mandatory = $true, ParameterSetName = 'Single', ValueFromPipelineByPropertyName)]
        [Alias('id')]
        [string[]]$MessageID,
        # Retrieve from folder.
        [Parameter(ParameterSetName = 'Single', ValueFromPipelineByPropertyName)]
        [Alias('parentFolderId')]
        [string]$Folder
    )
    try {
        $CommandName = $MyInvocation.InvocationName
        $RequestURI = [System.UriBuilder]::New('https', 'graph.microsoft.com')
        if ($Folder) {
            $RequestURI.Path = "v1.0/users/$($Mailbox)/mailfolders/$($Folder)/messages$($MessageID)"
        } else {
            $RequestURI.Path = "v1.0/users/$($Mailbox)/messages/$($MessageID)"
        }
        $DELETERequestParams = @{
            URI = $RequestURI.ToString()
            ContentType = 'application/json'
        }
        if ($PSCmdlet.ShouldProcess("Message $($MessageID)", 'Delete')) {
            $Result = New-MSGraphMailDELETERequest @DELETERequestParams
            if ($Result.StatusCode -eq 204) {
                Write-CustomMessage -Message "Removed message with ID $($MessageID)" -Type 'Success'
            }
        }
    } catch {
        $Command = $CommandName -Replace '-', ''
        $ErrorRecord = @{
            ExceptionType = 'System.Exception'
            ErrorMessage = "$($CommandName) failed."
            InnerException = $_.Exception
            ErrorID = "MicrosoftGraph$($Command)CommandFailed"
            ErrorCategory = 'ReadError'
            TargetObject = $_.TargetObject
            ErrorDetails = $_.ErrorDetails
            BubbleUpDetails = $True
        }
        $CommandError = New-MSGraphErrorRecord @ErrorRecord
        $PSCmdlet.ThrowTerminatingError($CommandError)
    }
}