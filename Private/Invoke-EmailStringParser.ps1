function Invoke-EmailStringParser {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [String[]]$Strings
    )
    # Split input string on ";" character.
    if ($Strings.Length -ge 2) {
        $EmailStrings = $Strings
    } else {
        $EmailStrings = $Strings.Split(";")
    }
    # Loop over each email string and add it to a hashtable in the expected format for an IMicrosoftGraphRecipient[] object.
    $EmailAddresses = foreach ($EmailString in $EmailStrings) {
        $ParsedEmailString = [regex]::Matches($EmailString, '\s?"?((?<name>.*?)"?\s*<)?(?<email>.*?[^>]*)')
        $Name = $ParsedEmailString[0].Groups['name'].value
        $Address = $ParsedEmailString[0].Groups['email'].value
        Write-Debug "Got name $($Name) and email $($Address) from string $($EmailString)"
        # Add the email address in the expected format for an IMicrosoftGraphEmailAddress object.
        $EmailAddress = @{
            'emailAddress' = @{
                name = $Name
                address = $Address
            }
        }
        $EmailAddress
    }
    return $EmailAddresses
}