function Send-MSGraphMail {
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
        [string]$Folder
    )
    try {
        $RequestURI = [System.UriBuilder]::New('https', 'graph.microsoft.com')
        if ($Folder) {
            $RequestURI.Path = "v1.0/users/$($Mailbox)/mailfolders/$($Folder)/messages/$($MessageID)/send"
        } else {
            $RequestURI.Path = "v1.0/users/$($Mailbox)/messages/$($MessageID)/send"
        }
        $POSTRequestParams = @{
            URI = $RequestURI.ToString()
            ContentType = 'application/json; charset=utf-8'
        }
        $Message = New-MSGraphMailPOSTRequest @POSTRequestParams
        Write-Debug "Microsoft Graph returned $($Message)"
        if ($Message) {
            Write-CustomMessage -Message "Sent message '$($Message.subject)' with ID $($Message.id)" -Type 'Success'
        }
    } catch {
        New-MSGraphError $_
    }
}