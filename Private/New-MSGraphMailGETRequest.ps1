function New-MSGraphMailGETRequest {
    <#
        .SYNOPSIS
            Builds a GET request for the Microsoft Graph API.
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
        # Use HTTPClient instead of Invoke-WebRequest.
        [switch]$UseHTTPClient
    )
    if ($null -eq $Script:MSGMConnectionInformation) {
        Throw "Missing Microsoft Graph connection information, please run 'Connect-MSGraphMail' first."
    }
    if ($null -eq $Script:MSGMAuthenticationInformation) {
        Throw "Missing Microsoft Graph authentication tokens, please run 'Connect-MSGraphMail' first."
    }
    try {
        if ($UseHTTPClient) {
            $WebRequestParams = @{
                Method = 'GET'
                URI = $URI
                ContentType = $ContentType
            }
            Write-Debug "Building new Microsoft Graph GET request with params: $($WebRequestParams | Out-String)"
            $Result = Invoke-MSGraphHTTPClientRequest @WebRequestParams
        } else {
            $WebRequestParams = @{
                Method = 'GET'
                URI = $URI
                ContentType = $ContentType
            }
            Write-Debug "Building new Microsoft Graph GET request with params: $($WebRequestParams | Out-String)"
            $Result = Invoke-MSGraphWebRequest @WebRequestParams
        }
        if ($Result) {
            Write-Debug "Microsoft Graph request returned $($Result | Out-String)"
            Return $Result
        } else {
            Throw 'Failed to process GET request.'
        }
    } catch {
        $ErrorRecord = @{
            ExceptionType = 'System.Net.Http.HttpRequestException'
            ErrorMessage = 'GET request sent to the Microsoft Graph API failed.'
            InnerException = $_.Exception
            ErrorID = 'MSGraphMailGetRequestFailed'
            ErrorCategory = 'ProtocolError'
            TargetObject = $_.TargetObject
            ErrorDetails = $_.ErrorDetails
            BubbleUpDetails = $True
        }
        $RequestError = New-MSGraphErrorRecord @ErrorRecord
        $PSCmdlet.ThrowTerminatingError($RequestError)
    }
}