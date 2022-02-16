function New-MSGraphMail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String[]]$From,
        [Parameter(Mandatory = $true)]
        [String[]]$To,
        [String[]]$CC,
        [String[]]$BCC,
        [String]$Subject,
        [Parameter(Mandatory = $true)]
        [String]$BodyContent,
        [String]$FooterContent,
        [Parameter(Mandatory = $true)]
        [string]$BodyFormat,
        [String]$Folder,
        [String[]]$Attachments,
        [String[]]$InlineAttachments,
        [Switch]$Draft,
        [Switch]$RequestDeliveryReceipt,
        [Switch]$RequestReadReceipt,
        [Switch]$Pipeline,
        [Switch]$Send,
        [Switch]$SaveandSend
    )
    try {
        $CommandName = $MyInvocation.InvocationName
        $MailFrom = Invoke-EmailStringParser -Strings $From
        $MailTo = Invoke-EmailStringParser -Strings @($To)
        if ($CC) {
            $MailCC = Invoke-EmailStringParser -Strings @($CC)
        } else {
            $MailCC = @()
        }
        if ($BCC) {
            $MailBCC = Invoke-EmailStringParser -Strings @($BCC)
        } else {
            $MailBCC = @()
        }
        $MailBody = New-MSGraphMailBody -BodyFormat $BodyFormat -BodyContent $BodyContent -FooterContent $FooterContent
        $MailParams = @{
            toRecipients = @($MailTo)
            from = $MailFrom
            subject = $Subject
            body = $MailBody
            ccRecipients = @($MailCC)
            bccRecipients = @($MailBCC)
        }
        if ($Draft) {
            $MailParams.isDraft = $true
        }
        if ($RequestDeliveryReceipt) {
            $MailParams.isDeliveryReceiptRequested = $true
        }
        if ($RequestReadReceipt) {
            $MailParams.isReadReceiptRequested = $true
        }
        $RequestURI = [System.UriBuilder]::New('https', 'graph.microsoft.com')
        if ($Folder) {
            $RequestURI.Path = "v1.0/users/$($MailFrom.EmailAddress.Address)/mailfolders/$($Folder)/messages"
        } elseif ($Send) {
            $RequestURI.Path = "v1.0/users/$($MailFrom.EmailAddress.Address)/sendmail"
        } else {
            $RequestURI.Path = "v1.0/users/$($MailFrom.EmailAddress.Address)/messages"
        }
        $POSTRequestParams = @{
            URI = $RequestURI.ToString()
            ContentType = 'application/json'
            Body = $MailParams
        }
        $Message = New-MSGraphMailPOSTRequest @POSTRequestParams
        Write-Debug "Microsoft Graph returned $($Message)"
        if ($Message) {
            Write-CustomMessage -Message "Created message '$($Message.subject)' with ID $($Message.id)" -Type 'Success'
        }
        if ($Attachments) {
            $AttachmentParams = @{
                Mailbox = $MailFrom.EmailAddress.Address
                MessageID = $Message.id
                Attachments = $Attachments
            }
            New-MSGraphMailAttachment @AttachmentParams | Out-Null
        }
        if ($InlineAttachments) {
            $InlineAttachmentParams = @{
                Mailbox = $MailFrom.EmailAddress.Address
                MessageID = $Message.id
                Attachments = $InlineAttachments
                InlineAttachments = $True
            }
            New-MSGraphMailAttachment @InlineAttachmentParams | Out-Null
        }
        if ($Pipeline -and $Message) {
            $Result = [PSCustomObject]@{
                id = $($Message).id
                mailbox = $MailFrom.EmailAddress.Address
                folder = $($Message).parentFolderId
            }
            Return $Result
        } elseif ($SaveandSend) {
            $SendParams = @{
                MessageID = $($Message).id
                Mailbox = $MailFrom.EmailAddress.Address
                Folder = $($Message).parentFolderId
            }
            Send-MSGraphMail @SendParams
        } elseif ($Message) {
            Return $Message
        }  
    } catch {
        $Command = $CommandName -Replace '-', ''
        $ErrorRecord = @{
            ExceptionType = 'System.Exception'
            ErrorMessage = "$($CommandName) failed."
            InnerException = $_.Exception
            ErrorID = "MicrosoftGraph$($Command)CommandFailed"
            ErrorCategory = 'ReadError'
            TargetObject = $_.TargetObject
            ErrorDetails = $_.ErrorDetails
            BubbleUpDetails = $True
        }
        $CommandError = New-MSGraphErrorRecord @ErrorRecord
        $PSCmdlet.ThrowTerminatingError($CommandError)
    }
}