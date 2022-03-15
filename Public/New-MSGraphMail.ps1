function New-MSGraphMail {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Does not change system state.')]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'MIME')]
        [String]$Mailbox,
        [Parameter(Mandatory = $true, ParameterSetName = 'Standard')]
        [String[]]$From,
        [Parameter(Mandatory = $true, ParameterSetName = 'Standard')]
        [String[]]$To,
        [Parameter(ParameterSetName = 'Standard')]
        [String[]]$CC,
        [Parameter(ParameterSetName = 'Standard')]
        [String[]]$BCC,
        [Parameter(ParameterSetName = 'Standard')]
        [String]$Subject,
        [Parameter(Mandatory = $true, ParameterSetName = 'Standard')]
        [String]$BodyContent,
        [Parameter(ParameterSetName = 'Standard')]
        [String]$FooterContent,
        [Parameter(Mandatory = $true, ParameterSetName = 'Standard')]
        [ValidateSet('HTML', 'text')]
        [string]$BodyFormat,
        [Parameter(Mandatory = $true, ParameterSetName = 'MIME')]
        [String]$MIMEMessage,
        [Parameter(ParameterSetName = 'Standard')]
        [Parameter(ParameterSetName = 'MIME')]
        [String]$Folder,
        [Parameter(ParameterSetName = 'Standard')]
        [String[]]$Attachments,
        [Parameter(ParameterSetName = 'Standard')]
        [String[]]$InlineAttachments,
        [Parameter(ParameterSetName = 'Standard')]
        [Switch]$Draft,
        [Parameter(ParameterSetName = 'Standard')]
        [Switch]$RequestDeliveryReceipt,
        [Parameter(ParameterSetName = 'Standard')]
        [Switch]$RequestReadReceipt,
        [Parameter(ParameterSetName = 'Standard')]
        [Parameter(ParameterSetName = 'MIME')]
        [Switch]$Pipeline,
        [Parameter(ParameterSetName = 'Standard')]
        [Parameter(ParameterSetName = 'MIME')]
        [Switch]$Send,
        [Parameter(ParameterSetName = 'Standard')]
        [Parameter(ParameterSetName = 'MIME')]
        [Switch]$SaveandSend
    )
    try {
        Write-Verbose "Using parameter set $($PSCmdlet.ParameterSetName)."
        if ($PSCmdlet.ParameterSetName -eq 'Standard') {
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
            if ($Draft) {
                $MailParams.isDraft = $true
            }
            if ($RequestDeliveryReceipt) {
                $MailParams.isDeliveryReceiptRequested = $true
            }
            if ($RequestReadReceipt) {
                $MailParams.isReadReceiptRequested = $true
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
            $ContentType = 'application/json; charset=utf-8'
        } elseif ($PSCmdlet.ParameterSetName -eq 'MIME') {
            $MailParams = $MIMEMessage
            $ContentType = 'text/plain'
        }
        $RequestURI = [System.UriBuilder]::New('https', 'graph.microsoft.com')
        if ($Folder) {
            $MessageBody = $MailParams
            if ($PSCmdlet.ParameterSetName -eq 'Standard') {
                $RequestURI.Path = "v1.0/users/$($MailFrom.EmailAddress.Address)/mailfolders/$($Folder)/messages"
            } elseif ($PSCmdlet.ParameterSetName -eq 'MIME') {
                $RequestURI.Path = "v1.0/users/$($Mailbox)/mailfolders/$($Folder)/messages"
            }
        } elseif ($Send) {
            if ($PSCmdlet.ParameterSetName -eq 'Standard') {
                $MessageBody = @{
                    message = $MailParams
                    saveToSentItems = $true
                }
                $RequestURI.Path = "v1.0/users/$($MailFrom.EmailAddress.Address)/sendmail"
            } elseif ($PSCmdlet.ParameterSetName -eq 'MIME') {
                $MessageBody = $MailParams
                $RequestURI.Path = "v1.0/users/$($Mailbox)/sendmail"
            }
            
        } else {
            $MessageBody = $MailParams
            if ($PSCmdlet.ParameterSetName -eq 'Standard') {
                $RequestURI.Path = "v1.0/users/$($MailFrom.EmailAddress.Address)/messages"
            } elseif ($PSCmdlet.ParameterSetName -eq 'MIME') {
                $RequestURI.Path = "v1.0/users/$($Mailbox)/messages"
            }
        }
        $POSTRequestParams = @{
            URI = $RequestURI.ToString()
            ContentType = $ContentType
            Body = $MessageBody
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
        New-MSGraphError $_
    }
}