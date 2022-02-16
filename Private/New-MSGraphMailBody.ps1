function New-MSGraphMailBody {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private function - no need to support.')]
    Param (
        [Parameter(Mandatory = $True)]
        [ValidateSet('text', 'html')]
        [string]$BodyFormat,
        [Parameter(Mandatory = $True)]
        [string]$BodyContent,
        [string]$FooterContent
    )
    if (Test-Path $BodyContent) {
        $MailContent = (Get-Content $BodyContent -Raw)
    } else {
        $MailContent = $BodyContent
        if (Test-Path $FooterContent) {
            $MailFooter = (Get-Content $FooterContent -Raw)
        } else { $MailFooter = $FooterContent }
        $MailBody = @{
            content     = "$($MailContent)$([System.Environment]::NewLine)$($MailFooter)"
            contentType = $BodyFormat
        }
        Return $MailBody
    }
}
