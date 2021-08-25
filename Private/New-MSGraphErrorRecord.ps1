using namespace System.Collections.Generic
function New-MSGraphErrorRecord {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private function - no need to support.')]
    param (
        [Parameter(Mandatory = $true)]
        [type]$ExceptionType,
        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage,
        [exception]$InnerException = $null,
        [Parameter(Mandatory = $true)]
        [string]$ErrorID,
        [Parameter(Mandatory = $true)]
        [errorcategory]$ErrorCategory,
        [object]$TargetObject = $null,
        [object]$ErrorDetails = $null,
        [switch]$BubbleUpDetails
    )
    $ExceptionMessage = [list[string]]::New()
    $ExceptionMessage.Add($ErrorMessage)
    if ($ErrorDetails.Message) {
        $MSGraphError = $_.ErrorDetails.Message | ConvertFrom-Json
        if ($MSGraphError.Message) {
            $ExceptionMessage.Add("The Microsoft Graph API said $($MSGraphError.ClassName): $($MSGraphError.Message).")
        }
    }
    if ($InnerException.Response) {
        $Response = $InnerException.Response
    }
    if ($InnerException.InnerException.Response) {
        $Response = $InnerException.InnerException.Response
    }
    if ($InnerException.InnerException.InnerException.Response) {
        $Response = $InnerException.InnerException.InnerException.Response
    }
    if ($Response) {
        $ExceptionMessage.Add("The Microsoft Graph API provided the status code $($Response.StatusCode.Value__): $($Response.ReasonPhrase).")
    }
    $ExceptionMessage.Add('You can use "Get-Error" for detailed error information.')
    $MSGraphError = [ErrorRecord]::New(
        $ExceptionType::New(
            $ExceptionMessage,
            $InnerException
        ),
        $ErrorID,
        $ErrorCategory,
        $TargetObject
    )
    if ($BubbleUpDetails) {
        $MSGraphError.ErrorDetails = $ErrorDetails
    }
    Return $MSGraphError
}