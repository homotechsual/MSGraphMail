#Requires -Version 7
function Invoke-MSGraphWebRequest {
    <#
        .SYNOPSIS
            Sends a request to the Microsoft Graph API using Invoke-WebRequest.
        .DESCRIPTION
            Wrapper function to send web requests to the Microsoft Graph API using the Invoke-WebRequest commandlet.
        .OUTPUTS
            Outputs an object containing the response from the web request.
    #>
    [Cmdletbinding()]
    [OutputType([Object])]
    param (
        # The request URI.
        [Parameter(Mandatory = $True)]
        [uri]$URI,
        # The request method.
        [Parameter(Mandatory = $True)]
        [string]$Method,
        # Don't authenticate.
        [switch]$Anonymous,
        # The content type for the request.
        [string]$ContentType,
        # The body content of the request.
        [object]$Body,
        # Additional headers.
        [hashtable]$AdditionalHeaders
    )
    $ProgressPreference = 'SilentlyContinue'
    if ([DateTime]::Now -ge $Script:MSGMAuthenticationInformation.Expires) {
        Write-Verbose 'The auth token has expired, renewing.'
        $ReconnectParameters = @{
            Reconnect = $True
        }
        Connect-MSGraphMail @ReconnectParameters
    }
    if ($null -ne $Script:MSGMAuthenticationInformation -and (-not $Anonymous)) {
        $AuthHeader = @{
            Authorization = "$($Script:MSGMAuthenticationInformation.Type) $($Script:MSGMAuthenticationInformation.Token)"
        }
    }
    if ($null -ne $AdditionalHeaders) {
        $RequestHeaders = $AuthHeader + $AdditionalHeaders
    } else {
        $RequestHeaders = $AuthHeader
    }
    if ($Method -eq 'PUT') {
        $SkipHeaderValidation = $True
    } else {
        $SkipHeaderValidation = $False
    }
    try {
        Write-Verbose "Making a $($Method) request to $($URI)"
        Write-Debug "Request headers: $($RequestHeaders | Out-String -Width 5000)"
        $WebRequestParams = @{
            URI = $URI
            Method = $Method
            ContentType = $ContentType
            Headers = $RequestHeaders
            SkipHeaderValidation = $SkipHeaderValidation
        }
        if ($Body -and (($Method -eq 'POST') -or ($Method -eq 'PUT'))) {
            $WebRequestParams.Body = $Body
        }
        $Response = Invoke-WebRequest @WebRequestParams
        Write-Debug "Response headers: $($Response.Headers | Out-String)"
        Write-Debug "Raw response: $($Response | Out-String -Width 5000)"
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
