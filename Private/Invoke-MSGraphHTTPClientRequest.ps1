using namespace System.Net.Http
using namespace System.Net.Http.Headers
#Requires -Version 7
function Invoke-MSGraphHTTPClientRequest {
    <#
        .SYNOPSIS
            Sends a request to the Microsoft Graph API using HTTPClient.
        .DESCRIPTION
            Wrapper function to send web requests to the Microsoft Graph API using the .NET HTTP client implementation.
        .OUTPUTS
            Outputs an object containing the response from the web request.
    #>
    [Cmdletbinding()]
    [OutputType([Object])]
    param (
        # The request URI.
        [Parameter(Mandatory = $True)]
        [uri]$URI,
        [Parameter(Mandatory = $True)]
        [string]$Method,
        [object]$Body,
        # The content type for the request.
        [string]$ContentType
    )
    $ProgressPreference = 'SilentlyContinue'
    if ([DateTime]::Now -ge $Script:MSGMAuthenticationInformation.Expires) {
        Write-Verbose 'The auth token has expired, renewing.'
        $ReconnectParameters = @{
            Reconnect = $True
        }
        Connect-MSGraphMail @ReconnectParameters
    }
    if (($null -ne $Script:MSGMAuthenticationInformation) -and ($Method -ne 'PUT')) {
        $AuthHeader = [AuthenticationHeaderValue]::New($Script:MSGMAuthenticationInformation.Type, $Script:MSGMAuthenticationInformation.Token)
    }
    try {
        Write-Verbose "Making a $($Method) request to $($URI)"
        Write-Debug "Authentication headers: $($AuthHeader.ToString())"        
        $HTTPClient = [HttpClient]::new()
        $HTTPClient.DefaultRequestHeaders.Authorization = $AuthHeader
        $HTTPClient.DefaultRequestHeaders.Add('Prefer', 'IdType%3D%22ImmutableId%22')
        if ($Method = 'GET') {
            $Request = $HTTPClient.GetAsync($URI)
        } elseif ($Method = 'PUT') {
            if (-Not $Body) {
                Throw 'Body is missing on PUT request.'
            }
            $Request = $HTTPClient.PutAsync($URI, $Body)
        }
        $Request.Wait()
        $Result = $Request.Result
        if ($Result.isFaulted) {
            Throw $Result.Exception
        }
        $Response = $Result.Content.ReadAsStringAsync().Result
        Write-Debug "Response headers: $($Result.Headers | Out-String)"
        Write-Debug "Raw response: $($Result | Out-String)"
        return $Response
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
