function New-MSGraphMailBody {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [ValidateSet('text','html')]
        [string]$BodyFormat,
        [Parameter(Mandatory=$True)]
        [string]$BodyContent,
        [string]$FooterContent
    )
    if (Test-Path $BodyContent) {
        $MailContent = (Get-Content $BodyContent -Raw)
    }
    if (Test-Path $FooterContent) {
        $MailFooter = (Get-Content $FooterContent -Raw)
    }
    $MailBody = @{
        content = "$($MailContent)$([System.Environment]::NewLine)$($MailFooter)"
        contentType = $BodyFormat
    }
    Return $MailBody
}