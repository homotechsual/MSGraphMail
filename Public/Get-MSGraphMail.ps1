function Get-MSGraphMail {
    [CmdletBinding()]
    param (
        # Specify the mailbox (or UPN) to retrieve emails for.
        [Parameter(Mandatory = $true, ParameterSetName = 'Multi')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Single', ValueFromPipelineByPropertyName)]
        [string]$Mailbox,
        # Retrieve a single message using a message ID.
        [Parameter(Mandatory = $true, ParameterSetName = 'Single', ValueFromPipelineByPropertyName)]
        [Alias('id')]
        [string[]]$MessageID,
        # Retrieve from folder.
        [Parameter(ParameterSetName = 'Multi')]
        [Parameter(ParameterSetName = 'Single', ValueFromPipelineByPropertyName)]
        [Alias('parentFolderId')]
        [string]$Folder,
        # Retrieve headers only.
        [Parameter(ParameterSetName = 'Multi')]
        [switch]$HeadersOnly,
        # Retrieve the message in MIME format.
        [Parameter(ParameterSetName = 'Single')]
        [switch]$MIME,
        # Search for emails based on a string.
        [Parameter(ParameterSetName = 'Multi')]
        [string]$Search,
        # Return this number of results.
        [Parameter(ParameterSetName = 'Multi')]
        [int]$PageSize = 500,
        # Transform the output into an object suitable for piping to other commands.
        [Parameter(ParameterSetName = 'Multi')]
        [Parameter(ParameterSetName = 'Single')]
        [switch]$Pipeline
    )
    try {
        $QueryStringCollection = [system.web.httputility]::ParseQueryString([string]::Empty)
        if ($HeadersOnly) {
            $QueryStringCollection.Add('$select', 'internetMessageHeaders')
        }
        if ($Search) {
            $QueryStringCollection.Add('$search', $Search)
        }
        if (($PageSize) -and ($PSCmdlet.ParameterSetName -ne 'Single')) {
            $QueryStringCollection.Add('$top', $PageSize)
        }
        $RequestURI = [System.UriBuilder]::New('https', 'graph.microsoft.com')
        if ($MessageID) {
            $RequestURI.Path = "v1.0/users/$($Mailbox)/messages/$($MessageID)"
            $ContentType = 'application/json'
            if ($MIME) {
                $RequestURI.Path = "v1.0/users/$($Mailbox)/messages/$($MessageID)/`$value"
                $ContentType = 'text/plain'
            }
        } elseif ($Folder) {
            $RequestURI.Path = "v1.0/users/$($Mailbox)/mailfolders/$($Folder)/messages"
        } else {
            $RequestURI.Path = "v1.0/users/$($Mailbox)/messages"
            $ContentType = 'application/json'
        }
        if ($QueryStringCollection.Count -gt 0) {
            $RequestURI.Query = $QueryStringCollection.toString()
        }
        $GETRequestParameters = @{
            URI = $RequestURI.ToString()
            ContentType = $ContentType
            UseHTTPClient = $True
        }
        $Content = New-MSGraphMailGETRequest @GETRequestParameters
        if ($Content) {
            if (-not $MIME) {
                $Content = $Content | ConvertFrom-Json
            } else {
                $Result = $Content
                Return $Result
            }
        }
        if ($Pipeline -and $Content) {
            $Result = [PSCustomObject]@{
                id = $($Content).value.id
                mailbox = $($Content).value.toRecipients.emailAddress.address
                folder = $($Content).value.parentFolderId
            }
            Return $Result
        } elseif ($Content.value) {
            $Result = $($Content.value)
            Return $Result
        }
    } catch {
        $ErrorRecord = @{
            ExceptionType = 'System.Net.Http.HttpRequestException'
            ErrorMessage = "Microsoft Graph API request $($_.TargetObject.Method) $($_.TargetObject.RequestUri) failed."
            InnerException = $_.Exception
            ErrorID = 'MicrosoftGraphRequestFailed'
            ErrorCategory = 'ProtocolError'
            TargetObject = $_.TargetObject
            ErrorDetails = $_.ErrorDetails
            BubbleUpDetails = $True
        }
        $RequestError = New-MSGraphErrorRecord @ErrorRecord
        $PSCmdlet.ThrowTerminatingError($RequestError)
    }
}