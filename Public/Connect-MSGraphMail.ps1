function Connect-MSGraphMail {
    [CmdletBinding()]
    Param(
        # Azure AD application id.
        [Parameter(Mandatory = $True, ParameterSetName = 'Connect')]
        [string]$ApplicationID,
        # Azure AD application secret.
        [Parameter(Mandatory = $True, ParameterSetName = 'Connect')]
        [string]$ApplicationSecret,
        # Graph permission scope.
        [Parameter(ParameterSetName = 'Connect')]
        [uri]$Scope = [uri]'https://graph.microsoft.com/.default',
        # Tenant ID.
        [Parameter(ParameterSetName = 'Connect')]
        [string]$TenantID,
        # Reconnect mode
        [Parameter(Mandatory = $True, ParameterSetName = 'Reconnect')]
        [switch]$Reconnect
    )
    if ((-not $Script:MSGMAuthenticationInformation.Token) -or ([DateTime]::Now -ge $Script:MSGMAuthenticationInformation.Expires)) {
        if (([DateTime]::Now -ge $Script:MSGMAuthenticationInformation.Expiry)) {
            try {
                if ((-not $Script:MSGMConnectionInformation) -and (-not $Reconnect)) {
                    $ConnectionInformation = @{
                        ClientID = $ApplicationID
                        ClientSecret = $ApplicationSecret
                        Scope = $Scope
                        URI = "https://login.microsoftonline.com/$($TenantID)/oauth2/v2.0/token"
                        TenantID = $TenantID
                    }
                    New-Variable -Name 'MSGMConnectionInformation' -Value $ConnectionInformation -Scope 'Script'
                }
                $AuthenticationBody = @{
                    client_id = $Script:MSGMConnectionInformation.ClientID
                    client_secret = $Script:MSGMConnectionInformation.ClientSecret
                    scope = $Script:MSGMConnectionInformation.Scope
                    grant_type = 'client_credentials'
                }
                $AuthenticationParameters = @{
                    URI = $Script:MSGMConnectionInformation.URI
                    Method = 'POST'
                    ContentType = 'application/x-www-form-urlencoded'
                    Body = $AuthenticationBody
                }
                $TokenResponse = Invoke-WebRequest @AuthenticationParameters
                $TokenPayload = ($TokenResponse.Content | ConvertFrom-Json)
                $AuthenticationInformation = @{
                    Token = $TokenPayload.access_token
                    Expires = Get-TokenExpiry -ExpiresIn $TokenPayload.expires_in
                    Type = $TokenPayload.token_type
                }
                if (-Not $Script:MSGMAuthenticationInformation) {
                    New-Variable -Name 'MSGMAuthenticationInformation' -Value $AuthenticationInformation -Scope 'Script'
                } else {
                    Set-Variable -Name 'MSGMAuthenticationInformation' -Value $AuthenticationInformation -Scope 'Script'
                }
                Write-CustomMessage -Message 'Connected to the Microsoft Graph API' -Type 'Success'
            } catch {
                $ErrorRecord = @{
                    ExceptionType = 'System.Net.Http.HttpRequestException'
                    ErrorMessage = "Graph API request $($_.TargetObject.Method) $($_.TargetObject.RequestUri) failed."
                    InnerException = $_.Exception
                    ErrorID = 'GraphAuthenticationFailed'
                    ErrorCategory = 'ProtocolError'
                    TargetObject = $_.TargetObject
                    ErrorDetails = $_.ErrorDetails
                    BubbleUpDetails = $True
                }
                $RequestError = New-HaloErrorRecord @ErrorRecord
                $PSCmdlet.ThrowTerminatingError($RequestError)
            }
        }
    } else {
        Write-CustomMessage -Message "Already connected to Microsoft Graph API." -Type 'Information'
    } 
}