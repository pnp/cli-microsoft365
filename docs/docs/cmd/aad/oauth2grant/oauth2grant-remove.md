# aad oauth2grant remove

Remove specified service principal OAuth2 permissions

## Usage

```sh
m365 aad oauth2grant remove [options]
```

## Options

`-h, --help`
: output usage information

`-i, --grantId <grantId>`
: `objectId` of OAuth2 permission grant to remove

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Before you can remove service principal's OAuth2 permissions, you need to get the `objectId` of the permissions grant to remove. You can retrieve it using the [aad oauth2grant list](./oauth2grant-list.md) command.

If the `objectId` listed when using the [aad oauth2grant list](./oauth2grant-list.md) command has a minus sign ('-') prefix, you may receive an error indicating `--grantId` is missing.  To resolve this issue simply escape the leading '-'.  

```sh
m365 aad oauth2grant remove --grantId \\-Zc1JRY8REeLxmXz5KtixAYU3Q6noCBPlhwGiX7pxmU
```

## Examples

Remove the OAuth2 permission grant with ID _YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek_

```sh
m365 aad oauth2grant remove --grantId YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek
```

## More information

- Application and service principal objects in Azure Active Directory (Azure AD): [https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects)