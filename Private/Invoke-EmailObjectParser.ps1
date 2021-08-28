function Invoke-EmailObjectParser {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Object[]]$Objects
    )
    # Loop over each email string and add it to a hashtable in the expected format for an IMicrosoftGraphRecipient[] object.
    Write-Debug "Email object parser received $($Objects | ConvertTo-Json)"
    $EmailAddresses = foreach ($EmailObject in $Objects) {
        $Name = $EmailObject.emailAddress.Name
        $Address = $EmailObject.emailAddress.Address
        Write-Debug "Got name $($Name) and email $($Address) from object $($EmailObject.emailAddress | Out-String)"
        # Turn the email into an output string.
        $EmailAddress = "$($Name) <$($Address)>"
        $EmailAddress
    }
    return $EmailAddresses -Join ';'
}