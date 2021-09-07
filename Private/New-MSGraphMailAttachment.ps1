function New-MSGraphMailAttachment {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [string]$Mailbox,
        [Parameter(Mandatory = $True)]
        [string]$MessageID,
        [string]$Folder,
        [Parameter(Mandatory = $True)]
        [string[]]$Attachments,
        [switch]$InlineAttachments
    )
    Write-Debug "Got attachments $($Attachments -join ', ')"
    foreach ($AttachmentItem in $Attachments) {
        if ($InlineAttachments) {
            $IAParts = $AttachmentItem.Split(';')
            $CID = $IAParts[0]
            Write-Verbose ("Content ID: $CID")
            $Attachment = $IAParts[1]
            Write-Verbose ("Attachment: $Attachment")
        } else {
            $Attachment = $AttachmentItem
        }
        Test-Path -Path $Attachment -ErrorAction Stop | Out-Null
        $AttachmentFile = Get-Item -Path $Attachment -ErrorAction Stop
        $Bytes = Get-Content -Path $AttachmentFile.FullName -AsByteStream -Raw
        if ($Bytes.Length -le 2999999) {
            Write-Debug "Attachment $($AttachmentFile.Fullname) size is $($Bytes.Length) which is less than 3MB - using direct upload"
            $UploadSession = $False
        } else {
            Write-Debug "Attachment $($AttachmentFile.Fullname) size is $($Bytes.Length) which is greater than 3MB - using streaming upload"
            $UploadSession = $True
        }
        $AttachmentItem = @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            attachmentType = 'file'
            name = $AttachmentFile.Name
        }
        if ($CID) {
            $AttachmentItem.contentId = $CID
            $AttachmentItem.isInline = $True
            if ($AttachmentFile.Extension -eq '.png') {
                $AttachmentItem.contentType = 'image/png'
            } elseif (($AttachmentFile.Extension -eq '.jpg') -or ($AttachmentFile.Extension -eq '.jpeg')) {
                $AttachmentItem.contentType = 'image/jpeg'
            } elseif ($AttachmentFile.Extension -eq '.gif') {
                $AttachmentItem.contentType = 'image/gif'
            }
        } else {
            $AttachmentItem.size = $($Bytes.Length)
        }
        Write-Debug "Generated attachment item $($AttachmentItem | ConvertTo-JSON)"
        $RequestURI = [System.UriBuilder]::New('https', 'graph.microsoft.com')
        if ($UploadSession) {
            $UploadTry = 0
            do {
                if ($Folder) {
                    $RequestURI.Path = "v1.0/users/$($Mailbox)/mailFolders/$($Folder)/messages/$($MessageID)/attachments/createUploadSession"
                } else {
                    $RequestURI.Path = "v1.0/users/$($Mailbox)/messages/$($MessageID)/attachments/createUploadSession"
                }
                $SessionAttachmentItem = @{
                    AttachmentItem = $AttachmentItem
                }
                $UploadSessionParams = @{
                    URI = $RequestURI.ToString()
                    Body = $SessionAttachmentItem
                    ContentType = 'application/json'
                    Raw = $False
                }
                try {
                    $UploadTry++
                    $InternalServerError = $False
                    Write-CustomMessage "Attempting to upload $($AttachmentFile.FullName) attempt number $($UploadTry)" -Type 'Information'
                    $AttachmentSession = New-MSGraphMailPOSTRequest @UploadSessionParams
                    Write-Debug "Got upload session details $($AttachmentSession)"
                    $AttachmentSessionURI = $AttachmentSession.uploadurl
                } catch {
                    $ErrorRecord = @{
                        ExceptionType = 'System.Net.Http.HttpRequestException'
                        ErrorMessage = 'Creating session for attachment upload to the Microsoft Graph API failed.'
                        InnerException = $_.Exception
                        ErrorID = 'MSGraphMailFailedToGetAttachmentUploadSession'
                        ErrorCategory = 'ProtocolError'
                        TargetObject = $_.TargetObject
                        ErrorDetails = $_.ErrorDetails
                        BubbleUpDetails = $True
                    }
                    $RequestError = New-MSGraphErrorRecord @ErrorRecord
                    $PSCmdlet.ThrowTerminatingError($RequestError)
                }
                if ($AttachmentSession) {
                    $AdditionalHeaders = @{
                        "Content-Range" = "bytes 0-$($Bytes.Length -1)/$($Bytes.Length)"
                    }
                    $AttachmentUploadParams =@{
                        URI = $AttachmentSessionURI
                        Body = $Bytes
                        Anonymous = $True
                        AdditionalHeaders = $AdditionalHeaders
                        Raw = $False
                    }
                    try {
                        $AttachmentUpload = New-MSGraphMailPUTRequest @AttachmentUploadParams
                        if ($AttachmentUpload) {
                            $InternalServerError = $False
                            Write-CustomMessage -Message "Attached file '$($AttachmentFile.FullName)' to message $($MessageID)" -Type 'Success'
                        }
                    } catch {
                        if ($_.Exception.InnerException.InnerException.Response.StatusCode.value__ -eq 500) {
                            Write-Warning "Attempt to upload '$($AttachmentFile.FullName)' failed. Retrying."
                            $InternalServerError = $True
                        } else {
                            $ErrorRecord = @{
                                ExceptionType = 'System.Net.Http.HttpRequestException'
                                ErrorMessage = "Sending attachment '$($AttachmentFile.Name)' to the Microsoft Graph API failed."
                                InnerException = $_.Exception
                                ErrorID = 'MSGraphMailAttachmentUploadFailed'
                                ErrorCategory = 'ProtocolError'
                                TargetObject = $_.TargetObject
                                ErrorDetails = $_.ErrorDetails
                                BubbleUpDetails = $True
                            }
                            $RequestError = New-MSGraphErrorRecord @ErrorRecord
                            $PSCmdlet.ThrowTerminatingError($RequestError)
                        }
                    }
                } 
            } while (($InternalServerError) -and ($UploadTry -le 5))
        } else {
            if ($Folder) {
                $RequestURI.Path = "v1.0/users/$($Mailbox)/mailFolders/$($Folder)/messages/$($MessageID)/attachments"
            } else {
                $RequestURI.Path = "v1.0/users/$($Mailbox)/messages/$($MessageID)/attachments"
            }
            $AttachmentItem.contentBytes = [convert]::ToBase64String($Bytes)
            $SimpleAttachmentParams = @{
                URI = $RequestURI.ToString()
                Body = $($AttachmentItem)
                ContentType = 'application/json'
                Raw = $False
            }
            try {
                $AttachmentUpload = New-MSGraphMailPOSTRequest @SimpleAttachmentParams
                if ($AttachmentUpload) {
                    Write-CustomMessage -Message "Attached file '$($AttachmentFile.Name)' to message $($MessageID)" -Type 'Success'
                }
            } catch {
                $ErrorRecord = @{
                    ExceptionType = 'System.Net.Http.HttpRequestException'
                    ErrorMessage = "Sending attachment '$($AttachmentFile.Name)' to the Microsoft Graph API failed."
                    InnerException = $_.Exception
                    ErrorID = 'MSGraphMailAttachmentUploadFailed'
                    ErrorCategory = 'ProtocolError'
                    TargetObject = $_.TargetObject
                    ErrorDetails = $_.ErrorDetails
                    BubbleUpDetails = $True
                }
                $RequestError = New-MSGraphErrorRecord @ErrorRecord
                $PSCmdlet.ThrowTerminatingError($RequestError)
            }
        }
    }
}