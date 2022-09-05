# outlook mail send

Sends an email

## Usage

```sh
m365 outlook mail send [options]
```

## Options

`-s, --subject <subject>`
: Email subject

`-t, --to <to>`
: Comma-separated list of emails to send the message to.

`--cc [cc]`
: Comma-separated list of CC recipients for the message.

`--bcc [bcc]`
: Comma-separated list of BCC recipients for the message.

`--sender [sender]`
: Optional upn or user id to specify what account to send the message from. Also see the remarks section.

`-m, --mailbox [mailbox]`
: Specify this option to send the email on behalf of another mailbox, for example a shared mailbox, group or distribution list. The sender needs to be a delegate on the specified mailbox. Also see the remarks section.

`--bodyContents <bodyContents>`
: String containing the body of the email to send.

`--bodyContentType [bodyContentType]`
: Type of the body content. Available options: `Text,HTML`. Default `Text`.

`--importance [importance]`
: The importance of the message. Available options: `low`, `normal` or `high`. Default is `normal`.

`--saveToSentItems [saveToSentItems]`
: Save email in the sent items folder. Default `true`.

--8<-- "docs/cmd/_global.md"

## Remarks

### If you are connected using app only authentication

- Always specify a user id or upn in the `--sender` option. The email will be sent as if it came from the specified user, and can optionally be saved in the sent folder of that user account.
- You can optionally also specify the `--mailbox` option to send mail on behalf of a shared mailbox, group or distribution list. The account used in the `--sender` option, needs to have 'Send on behalf of' permissions on the mailbox in question.

!!! important
    You need `Mail.Send` application permissions on the Microsoft Graph to be able to send mails using an application identity. 

### If you are connected with a regular user account

- Specify the `--mailbox` option if you want to send an email on behalf of another mailbox. This can be a shared mailbox, group or distribution list. It will be visible in the email that the email was sent by you. You need to be assigned `Send on behalf of` permissions on the mailbox in question.  
- You can specify the `--sender` option if you want to send an email as if you were the other user.
The sent email can optionally be saved in the sent folder of that user account. You'll need `Read and manage (Full Access)` permissions on the mailbox of the other user. You can combine the `--sender` and `--mailbox` options to let the other user send a mail on behalf of the specified mailbox.

!!! important
    You need at least `Mail.Send` delegated permissions on the Microsoft Graph to be able to send emails. When specifying another user as sender, you'll need `Mail.Send.Shared` permissions.

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

Send an email on behalf of a shared mailbox using the signed in user

```sh
m365 outlook mail send --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the team site" --mailbox sales@contoso.com
```

Send an email as another user

```sh
m365 outlook mail send --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the team site" --sender svc_project@contoso.com
```

Send an email as another user, on behalf of a shared mailbox

```sh
m365 outlook mail send --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the team site" --sender svc_project@contoso.com --mailbox sales@contoso.com
```

Send an email with cc and bcc recipients marked as important

```sh
m365 outlook mail send --to chris@contoso.com --cc claire@contoso.com --bcc randy@contoso.com --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the team site" --importance high
```
