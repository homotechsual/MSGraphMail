function New-MSGraphMailPOSTRequest {
    <#
        .SYNOPSIS
            Builds a POST request for the Microsoft Graph API.
        .DESCRIPTION
            Wrapper function to build web requests for the Microsoft Graph API.
        .OUTPUTS
            Outputs an object containing the response from the web request.
    #>
    [CmdletBinding()]
    [OutputType([Object])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private function - no need to support.')]
    param (
        # The request URI.
        [uri]$URI,
        # The content type for the request.
        [string]$ContentType,
        # The request body.
        [object]$Body,
        # Don't authenticate.
        [switch]$Anonymous,
        # Additional headers.
        [hashtable]$AdditionalHeaders = $null,
        # Return raw result?
        [switch]$Raw
    )
    if ($null -eq $Script:MSGMConnectionInformation) {
        Throw "Missing Microsoft Graph connection information, please run 'Connect-MSGraphMail' first."
    }
    if ($null -eq $Script:MSGMAuthenticationInformation) {
        Throw "Missing Microsoft Graph authentication tokens, please run 'Connect-MSGraphMail' first."
    }
    try {
        $WebRequestParams = @{
            Method = 'POST'
            Uri = $URI
            ContentType = $ContentType
            Anonymous = $Anonymous
            Body = ($Body | ConvertTo-JSON -Depth 15)
            AdditionalHeaders = $AdditionalHeaders
        }
        #Write-Debug "Building new Microsoft Graph POST request with body: $($WebRequestParams | Out-String -Width 5000)"
        $Result = Invoke-MSGraphWebRequest @WebRequestParams
        if ($Result) {
            if ($Raw) {
                Return $Result
            } else {
                Return $Result.content | ConvertFrom-Json -Depth 5
            }
        } else {
            Throw 'No response to POST request'
        }
    } catch {
        $ErrorRecord = @{
            ExceptionType = 'System.Net.Http.HttpRequestException'
            ErrorMessage = 'POST request sent to the Microsoft Graph API failed.'
            InnerException = $_.Exception
            ErrorID = 'MSGraphMailPostRequestFailed'
            ErrorCategory = 'ProtocolError'
            TargetObject = $_.TargetObject
            ErrorDetails = $_.ErrorDetails
            BubbleUpDetails = $True
        }
        $RequestError = New-MSGraphErrorRecord @ErrorRecord
        $PSCmdlet.ThrowTerminatingError($RequestError)
    }
}