# outlook sendmail

Sends an email on behalf of the current user

## Usage

```sh
m365 outlook mail send [options]
```

## Alias

```sh
m365 outlook sendmail [options]
```

## Options

`-s, --subject <subject>`
: Email subject

`-t, --to <to>`
: Comma-separated list of emails to send the message to

`--bodyContents <bodyContents>`
: String containing the body of the email to send

`--bodyContentType [bodyContentType]`
: Type of the body content. Available options: `Text,HTML`. Default `Text`

`--saveToSentItems [saveToSentItems]`
: Save email in the sent items folder. Default `true`

`--attachment [attachment]`
: Path to the file that will be added as attachment to the email. Use this option multiple times for multiple attachments.

--8<-- "docs/cmd/_global.md"

## Remarks

When using the `attachment` option, note that the total size of all attachment files cannot exceed 3 MB size.

## Examples

Send a text email to the specified email address

```sh
m365 outlook mail send --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the team site"
```

Send an HTML email to the specified email addresses

```sh
m365 outlook mail send --to "chris@contoso.com,brian@contoso.com" --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the <a href='https://contoso.sharepoint.com/sites/marketing'>team site</a>" --bodyContentType HTML
```

Send an HTML email to the specified email address loading email contents from a file on disk

```sh
m365 outlook mail send --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContents @email.html --bodyContentType HTML
```

Send a text email to the specified email address. Don't store the email in sent items

```sh
m365 outlook mail send --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the team site" --saveToSentItems false
```

Send an email with attachments

```sh
m365 outlook mail send --to chris@contoso.com --subject "Monthly reports" --bodyContents "Here are the reports of this month." --attachment "C:/Reports/File1.jpg" --attachment "C:/Reports/File2.docx" --attachment "C:/Reports/File3.xlsx"
```
