# spo mail send

Sends an e-mail from SharePoint

## Usage

```sh
m365 spo mail send [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: Absolute URL of the site from which the email will be sent

`--to <to>`
: Comma-separated list of recipients' e-mail addresses

`--subject <subject>`
: Subject of the e-mail

`--body <body>`
: Content of the e-mail

`--from [from]`
: Sender's e-mail address

`--cc [cc]`
: Comma-separated list of CC recipients

`--bcc [bcc]`
: Comma-separated list of BCC recipients

`--additionalHeaders [additionalHeaders]`
: JSON string with additional headers

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

All recipients (internal and external) have to have access to the target SharePoint site.

## Examples

Send an e-mail to _user@contoso.com_

```sh
m365 spo mail send --webUrl https://contoso.sharepoint.com/sites/project-x --to "user@contoso.com" --subject "Email sent via CLI for Microsoft 365" --body "<h1>CLI for Microsoft 365</h1>Email sent via <b>command</b>."
```

Send an e-mail to multiples addresses

```sh
m365 spo mail send --webUrl https://contoso.sharepoint.com/sites/project-x --to "user1@contoso.com,user2@contoso.com" --subject "Email sent via CLI for Microsoft 365" --body "<h1>CLI for Microsoft 365</h1>Email sent via <b>command</b>." --cc "user3@contoso.com" --bcc "user4@contoso.com"
```

Send an e-mail to _user@contoso.com_ with additional headers

```sh
m365 spo mail send --webUrl https://contoso.sharepoint.com/sites/project-x --to "user@contoso.com" --subject "Email sent via CLI for Microsoft 365" --body "<h1>CLI for Microsoft 365</h1>Email sent via <b>command</b>." --additionalHeaders "'{\"X-MC-Tags\":\"CLI for Microsoft 365\"}'"
```
