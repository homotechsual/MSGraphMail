function Invoke-EmailObjectParser {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [Object[]]$Objects
    )
    # Loop over each email object and add it to a string.
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