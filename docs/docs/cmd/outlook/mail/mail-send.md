# outlook sendmail

Sends e-mail on behalf of the current user

## Usage

```sh
m365 outlook mail send [options]
```

## Alias

```sh
m365 outlook sendmail [options]
```

## Options

`-h, --help`
: output usage information

`-s, --subject <subject>`
: E-mail subject

`-t, --to <to>`
: Comma-separated list of e-mails to send the message to

`--bodyContents [bodyContents]`
: String containing the body of the e-mail to send

`--bodyContentsFilePath [bodyContentsFilePath]`
: Relative or absolute path to the file with e-mail body contents

`--bodyContentType [bodyContentType]`
: Type of the body content. Available options: `Text,HTML`. Default `Text`

`--saveToSentItems [saveToSentItems]`
: Save e-mail in the sent items folder. Default `true`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Send a text e-mail to the specified e-mail address

```sh
m365 outlook mail send --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the team site"
```

Send an HTML e-mail to the specified e-mail addresses

```sh
m365 outlook mail send --to "chris@contoso.com,brian@contoso.com" --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the <a href='https://contoso.sharepoint.com/sites/marketing'>team site</a>" --bodyContentType HTML
```

Send an HTML e-mail to the specified e-mail address loading e-mail contents from a file on disk

```sh
m365 outlook mail send --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContentsFilePath email.html --bodyContentType HTML
```

Send a text e-mail to the specified e-mail address. Don't store the e-mail in sent items

```sh
m365 outlook mail send --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the team site" --saveToSentItems false
```
