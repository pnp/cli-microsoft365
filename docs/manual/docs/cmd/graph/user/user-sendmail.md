# graph user sendmail

Sends e-mail on behalf of the current user

## Usage

```sh
graph user sendmail [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-s, --subject <subject>`|E-mail subject
`-t, --to <to>`|Comma-separated list of e-mails to send the message to
`--bodyContents [bodyContents]`|String containing the body of the e-mail to send
`--bodyContentsFilePath [bodyContentsFilePath]`|Relative or absolute path to the file with e-mail body contents
`--bodyContentType [bodyContentType]`|Type of the body content. Available options: `Text|HTML`. Default `Text`
`--saveToSentItems [saveToSentItems]`|Save e-mail in the sent items folder. Default `true`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To send an e-mail on behalf of the current user, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Send a text e-mail to the specified e-mail address

```sh
graph user sendmail --to chris@contoso.com --subject 'DG2000 Data Sheets' --bodyContents 'The latest data sheets are in the team site'
```

Send an HTML e-mail to the specified e-mail addresses

```sh
graph user sendmail --to chris@contoso.com,brian@contoso.com --subject 'DG2000 Data Sheets' --bodyContents 'The latest data sheets are in the <a href="https://contoso.sharepoint.com/sites/marketing">team site</a>' --bodyContentType HTML
```

Send an HTML e-mail to the specified e-mail address loading e-mail contents from a file on disk

```sh
graph user sendmail --to chris@contoso.com --subject 'DG2000 Data Sheets' --bodyContentsFilePath email.html --bodyContentType HTML
```

Send a text e-mail to the specified e-mail address. Don't store the e-mail in sent items

```sh
graph user sendmail --to chris@contoso.com --subject 'DG2000 Data Sheets' --bodyContents 'The latest data sheets are in the team site' --saveToSentItems false
```