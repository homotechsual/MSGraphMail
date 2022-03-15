function New-MSGraphMailBody {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private function - no need to support.')]
    Param (
        [Parameter(Mandatory = $True)]
        [ValidateSet('HTML', 'text')]
        [string]$BodyFormat,
        [Parameter(Mandatory = $True)]
        [string]$BodyContent,
        [string]$FooterContent
    )
    if (Test-Path $BodyContent) {
        $MailContent = (Get-Content $BodyContent -Raw)
        Write-Verbose "Using file $BodyContent as body content."
        Write-Debug "Body content: `r`n$MailContent"
    } else {
        $MailContent = $BodyContent
        Write-Verbose "Using string as body content."
        Write-Debug "Body content: `r`n$MailContent"
    }
    if (Test-Path $FooterContent) {
        $MailFooter = (Get-Content $FooterContent -Raw)
        Write-Verbose "Using file $FooterContent as footer content."
        Write-Debug "Footer content: `r`n$MailFooter"
    } else {
        $MailFooter = $FooterContent
        Write-Verbose "Using string as footer content."
        Write-Debug "Footer content: `r`n$MailFooter"
    }
    $MailBody = @{
        content     = "$($MailContent)$([System.Environment]::NewLine)$($MailFooter)"
        contentType = $BodyFormat
    }
    Return $MailBody
}
