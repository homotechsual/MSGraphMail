function Move-MSGraphMail {
    [CmdletBinding()]
    param (
        # Specify the mailbox (or UPN) to move emails for.
        [Parameter(Mandatory = $true, ParameterSetName = 'Single', ValueFromPipelineByPropertyName)]
        [string]$Mailbox,
        # Retrieve a single message using a message ID.
        [Parameter(Mandatory = $true, ParameterSetName = 'Single', ValueFromPipelineByPropertyName)]
        [Alias('id')]
        [string[]]$MessageID,
        # Retrieve from folder.
        [Parameter(ParameterSetName = 'Single', ValueFromPipelineByPropertyName)]
        [Alias('parentFolderId')]
        [string]$Folder,
        # Destination.
        [string]$Destination = 'deleteditems'
    )
    try {
        $CommandName = $MyInvocation.InvocationName
        $MoveParams = @{
            destinationId = $Destination
        }
        $RequestURI = [System.UriBuilder]::New('https', 'graph.microsoft.com')
        if ($Folder) {
            $RequestURI.Path = "v1.0/users/$($Mailbox)/mailfolders/$($Folder)/messages/$($MessageID)/move"
        } else {
            $RequestURI.Path = "v1.0/users/$($Mailbox)/messages/$($MessageID)/move"
        }
        $POSTRequestParams = @{
            URI = $RequestURI.ToString()
            ContentType = 'application/json'
            Body = $MoveParams
        }
        $Message = New-MSGraphMailPOSTRequest @POSTRequestParams
        Write-Debug "Microsoft Graph returned $($Message)"
        if ($Message) {
            Write-CustomMessage -Message "Moved message '$($Message.subject)' with ID $($Message.id) to folder $($Message.parentFolderId)" -Type 'Success'
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