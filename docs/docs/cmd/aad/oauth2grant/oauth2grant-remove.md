# aad oauth2grant remove

Remove specified service principal OAuth2 permissions

## Usage

```sh
m365 aad oauth2grant remove [options]
```

## Options

`-i, --grantId <grantId>`
: `objectId` of OAuth2 permission grant to remove

`--confirm`
: Do not prompt for confirmation before removing OAuth2 permission grant

--8<-- "docs/cmd/_global.md"

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

Remove the OAuth2 permission grant with ID _YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek_ without being asked for confirmation

```sh
m365 aad oauth2grant remove --grantId YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek --confirm
```

## More information

- Application and service principal objects in Azure Active Directory (Azure AD): [https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects)
- Delete a delegated permission grant (oAuth2PermissionGrant): [https://docs.microsoft.com/en-us/graph/api/oauth2permissiongrant-delete?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/oauth2permissiongrant-delete?view=graph-rest-1.0)