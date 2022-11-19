# aad user guest add

Invite an external user to the organization

## Usage

```sh
m365 aad user guest add [options]
```

## Options

`--emailAddress <emailAddress>`
: The email address of the user.

`--displayName [displayName]`
: The display name of the user.

`--inviteRedirectUrl [inviteRedirectUrl]`
: The URL the user should be redirected to once the invitation is redeemed. If not specified, default URL https://myapplications.microsoft.com will be set.

`--welcomeMessage [welcomeMessage]`
: Personal welcome message which will be added to the email along with the default email.

`--ccRecipients [ccRecipients]`
: Additional recipients the invitation message should be sent to. Currently only 1 additional recipient is supported.

`--messageLanguage [messageLanguage]`
: The language you want to send the default message in. The language format should be in [ISO 639](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/ed06cf15-306c-43be-9053-ca81ca51e656). The default is `en-US`.

`--sendInvitationMessage`
: Indicates whether an email should be sent to the user.

--8<-- "docs/cmd/_global.md"

## Examples

Invite a user via email and set the display name

```sh
m365 aad user guest add --emailAddress john.doe@contoso.com --displayName "John Doe" --sendInvitationMessage
```

Invite a user with a custom email and custom redirect url

```sh
m365 aad user guest add --emailAddress john.doe@contoso.com --welcomeMessage "Hi John, welcome to the organization!" --inviteRedirectUrl https://contoso.sharepoint.com --sendInvitationMessage
```

Invite a user and send an invitation mail in Dutch

```sh
m365 aad user guest add --emailAddress john.doe@contoso.com --messageLanguage nl-BE --sendInvitationMessage
```

## Response

=== "JSON"

    ```json
    {
      "id": "35f7f726-c541-4aef-a64e-a7b6868fe47f",
      "inviteRedeemUrl": "https://login.microsoftonline.com/redeem?rd=https%3a%2f%2finvitations.microsoft.com%2fredeem%2f%3ftenant%3db373bc30-03b3-49bc-be72-9dd3e9027da8%26user%3d35f7f726-c541-4aef-a64e-a7b6868fe47f%26ticket%3dCjO3u3ZpQF2uthfZETfZ8gURzod5egvYI0uhaSN1Loo%25253d%26ver%3d2.0",
      "invitedUserDisplayName": "John Doe",
      "invitedUserType": "Guest",
      "invitedUserEmailAddress": "john.doe@contoso.com",
      "sendInvitationMessage": true,
      "resetRedemption": false,
      "inviteRedirectUrl": "https://myapplications.microsoft.com/",
      "status": "PendingAcceptance",
      "invitedUserMessageInfo": {
        "messageLanguage": "en-US",
        "customizedMessageBody": "Hi John, welcome to the organization!",
        "ccRecipients": [
          {
            "emailAddress": {
              "name": null,
              "address": "maria.jones@contoso.com"
            }
          }
        ]
      },
      "invitedUser": {
        "id": "5257b5b2-4056-4a45-a05e-df5c92d53e6e"
      }
    }
    ```

=== "Text"

    ```text
    id                     : 35f7f726-c541-4aef-a64e-a7b6868fe47f
    inviteRedeemUrl        : https://login.microsoftonline.com/redeem?rd=https%3a%2f%2finvitations.microsoft.com%2fredeem%2f%3ftenant%3db373bc30-03b3-49bc-be72-9dd3e9027da8%26user%3d35f7f726-c541-4aef-a64e-a7b6868fe47f%26ticket%3dCjO3u3ZpQF2uthfZETfZ8gURzod5egvYI0uhaSN1Loo%25253d%26ver%3d2.0
    invitedUserDisplayName : John Doe
    invitedUserEmailAddress: liwidit556@adroh.com
    invitedUserType        : Guest
    resetRedemption        : false
    sendInvitationMessage  : true
    status                 : PendingAcceptance
    ```

=== "CSV"

    ```csv
    id,inviteRedeemUrl,invitedUserDisplayName,invitedUserEmailAddress,invitedUserType,resetRedemption,sendInvitationMessage,status
    35f7f726-c541-4aef-a64e-a7b6868fe47f,https://login.microsoftonline.com/redeem?rd=https%3a%2f%2finvitations.microsoft.com%2fredeem%2f%3ftenant%3db373bc30-03b3-49bc-be72-9dd3e9027da8%26user%3d35f7f726-c541-4aef-a64e-a7b6868fe47f%26ticket%3dCjO3u3ZpQF2uthfZETfZ8gURzod5egvYI0uhaSN1Loo%25253d%26ver%3d2.0,John Doe,liwidit556@adroh.com,Guest,,1,PendingAcceptance
    ```
