# MS Graph Mail - A pure-PowerShell Graph API mail client

![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/MSGraphMail?style=for-the-badge) ![PowerShell Gallery](https://img.shields.io/powershellgallery/dt/MSGraphMail?style=for-the-badge)

## Preparations

You will need the following:

1. An [**Azure AD Application ID**](https://aad.portal.azure.com)
2. An **Azure AD Application Secret**
3. An **Azure AD Tenant ID**
4. [**PowerShell 7**](https://aka.ms/powershell-release?tag=stable) installed on your Windows, Linux or MacOS device.
5. The code from this module in your PowerShell modules folder find this by running `$env:PSModulePath` in your PowerShell 7 session. Install from PSGallery with `Install-Module MSGraphMail`

## Import the Module

Run `Import-Module 'MSGraphMail'` to load the module into your current session.

## Connecting to the Microsoft Graph API

Connecting to the Microsoft Graph API uses the Azure AD Application information and the Connect-MSGraphMail client application.

Using the **Splatting** technique:

Splatting is a system in PowerShell that lets us put our parameters in a nicely formatted easy to read object (a HashTable to be specific!) and then "splat" them at the command. To do this, first things first setup a PowerShell object to hold your credentials. For example:

```powershell
$MSGraphMailConnectionParameters = @{
    ApplicationID = '<YOUR APPLICATION ID>'
    ApplicationSecret = '<YOUR APPLICATION SECRET>'
    TenantID = '<YOUR TENANT ID>'
}
Connect-MSGraphMail @MSGraphMailConnectionParameters
```

Using the **Traditional** technique:

If you don't want to - or can't "splat" - we can fall back on a more traditional route:

```powershell
Connect-MSGraphMail -ApplicationID '<YOUR APPLICATION ID>' -ApplicationSecret '<YOUR APPLICATION SECRET>' -TenantID '<YOUR TENANT ID>'
```

## Getting Emails

Getting emails hinges around a single command `Get-MSGraphMail` at it's most basic this looks like this.

Using the **Splatting** technique:

```powershell
$MailParameters = @{
    Mailbox = 'you@example.uk'
}
Get-MSGraphMail @MailParameters
```

Using the **Traditional** technique:

```powershell
Get-MSGraphMail -Mailbox 'you@example.uk'
```

You can get more specific with the following parameters:

* **MessageID** - Retrieves a single message by ID.
* **Folder** - Retrieves messages (or a single message) from a specific folder.
* **HeadersOnly** - Retrieves only the message headers.
* **MIME** - Retrieves a single message in MIME format (Requires **MessageID**).
* **Search** - Searches emails based on a string.
* **PageSize** - Retrieves only the given number of results.
* **Pipeline** - Formats the output for Pipelining to other commands - like `Move-MSGraphMail` or `Delete-MSGraphMail`.
* **Select** - Retrieves only the specified fields from the Graph API.
* **Summary** - Displays a summary of the message(s) retrieved. See #1 for details.

## Creating an E-Mail

Creating an email requires passing parameters to the `New-MSGraphMail` commandlet like so:

Using the **Splatting** technique:

```powershell
$MailParameters = @{
    From = 'You <you@example.uk>'
    To = 'Them <them@example.com>', 'Someone <someone@example.com>'
    Subject = 'Your invoice #1234 is ready.'
    BodyContent = 'X:\Emails\BodyContent.txt'
    FooterContent = 'X:\Emails\FooterContent.txt'
    Attachments = 'X:\Files\SendtoExample.docx','X:\Files\SendToExample.zip'
    BodyFormat = 'text'
}
New-MSGraphMail @MailParameters
```

Using the **Traditional** technique:

```powershell
New-MSGraphMail -From 'You <you@example.uk>' -To 'Them <them@example.com>', 'Someone <someone@example.com>' -Subject 'Your invoice #1234 is ready.' -BodyContent 'X:\Emails\BodyContent.txt' -FooterContent 'X:\Emails\FooterContent.txt' -Attachments 'X:\Files\SendtoExample.docx','X:\Files\SendToExample.zip' -BodyFormat 'text'
```

If this works we'll see:

> SUCCESS: Created message 'Your invoice #1234 is ready.' with ID AAMkADg0MTI1YTY5LTZhNTAtNGY2Ni1iYmFmLTYyNTIxNmQ3ZTAyMQBGAAAAAADcjV4oGXn1Sb6mQOgHYL6tBwAynr9oS8bwR42_Ec20-qUkAAAAAAEQAAAynr9oS8bwR42_Ec20-qUkAAcuZgfeAAA=

A draft email will have appeared in the account provided to `From`. Unless you specify the `-Send` parameter which immediately sends the email bypassing the draft creation.

You can use inline attachments by using `-InlineAttachments` and specifying attachments in the format `'cid;filepath'` e.g:

```powershell
New-MSGraphMail -From 'You <you@example.uk>' -To 'Them <them@example.com>', 'Someone <someone@example.com>' -Subject 'Your invoice #1234 is ready.' -BodyContent 'X:\Emails\BodyContent.html' -FooterContent 'X:\Emails\FooterContent.html' -Attachments 'X:\Files\SendtoExample.docx','X:\Files\SendToExample.zip' -BodyFormat 'html' -InlineAttachments 'signaturelogo;X\Common\EmailSignatureLogo.png', 'productlogo;X:\Products\Widgetiser\WidgetiserLogoEmail.png'
```

The two inline attachments would map to:

```html
<img alt="Our Logo" src="cid:signaturelogo"/>
```

and

```html
<img alt="Widgetiser Logo" src="cid:productlogo"/>
```

respectively.

## Sending an E-Mail

Sending an email requires one small alteration of the above command - adding:

```powershell
Pipeline = $True
```

if splatting or

```powershell
-Pipeline
```

if using traditional parameter passing.

This tells the command that we're going to pipeline the output - specifically that we're going to send it to another command. In our case we'd end up doing:

```powershell
New-MSGraphMail @MailParameters | Send-MSGraphMail
```

The important part here is `| Send-MSGraphMail` quite literally `|` or **Pipe** and then the next command.

## Moving an E-Mail

Moving an email requires one small alteration of the Get or New command - adding:

```powershell
Pipeline = $True
```

if splatting or

```powershell
-Pipeline
```

if using traditional parameter passing.

This tells the command that we're going to pipeline the output - specifically that we're going to send it to another command. In our case we'd end up doing:

```powershell
New-MSGraphMail @MailParameters | Move-MSGraphMail -Destination 'deleteditems'
```

The `-Destination` parameter for `Move-MSGraphMail` accepts a "well known folder name" e.g: "deleteditems" or "drafts" or "inbox" or a Folder ID.

The important part here is `| Move-MSGraphMail` quite literally **Pipe (|)** and then the next command.

## Deleting an E-Mail

If you want to "permanently" delete an email you can pipe the email to the `Remove-MSGraphMail` command. Similar to moving an email this works as so:

```powershell
Get-MSGraphMail @MailParameters | Remove-MSGraphMail -Confirm:$False
```

Disecting this - we're getting the mail and then passing it down the pipeline an telling `Remove-MSGraphMail` not to prompt us for permission by setting `-Confirm:$False`
