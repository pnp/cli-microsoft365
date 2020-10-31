# aad oauth2grant set

Update OAuth2 permissions for the service principal

## Usage

```sh
m365 aad oauth2grant set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --grantId <grantId>`
: `objectId` of OAuth2 permission grant to update

`-s, --scope <scope>`
: Permissions to grant

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Before you can update service principal's OAuth2 permissions, you need to get the `objectId` of the permissions grant to update. You can retrieve it using the [aad oauth2grant list](./oauth2grant-list.md) command.

If the `objectId` listed when using the [aad oauth2grant list](./oauth2grant-list.md) command has a minus sign ('-') prefix, you may receive an error indicating `--grantId` is missing.  To resolve this issue simply escape the leading '-'.  

```sh
m365 aad oauth2grant set --grantId \\-Zc1JRY8REeLxmXz5KtixAYU3Q6noCBPlhwGiX7pxmU
```

## Examples

Update the existing OAuth2 permission grant with ID _YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek_ to the _Calendars.Read Mail.Read_ permissions

```sh
m365 aad oauth2grant set --grantId YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek --scope "Calendars.Read Mail.Read"
```

## More information

- Application and service principal objects in Azure Active Directory (Azure AD): [https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects)