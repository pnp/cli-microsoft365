# aad app role delete

Deletes role from the specified Azure AD app registration

## Usage

```sh
m365 aad app role delete [options]
```

## Options

`--appId [appId]`
: Application (client) ID of the Azure AD application registration from which role should be deleted. Specify either `appId`, `appObjectId` or `appName`

`--appObjectId [appObjectId]`
: Object ID of the Azure AD application registration from which role should be deleted. Specify either `appId`, `appObjectId` or `appName`

`--appName [appName]`
: Name of the Azure AD application registration from which role should be deleted. Specify either `appId`, `appObjectId` or `appName`

`-n, --name [name]`
: Name of the role to delete. Specify either `name`, `id` or `claim`

`-i, --id [id]`
: Id of the role to delete. Specify either `name`, `id` or `claim`

`-c, --claim [claim]`
: Claim value of the role to delete. Specify either `name`, `id` or `claim`

`--confirm`
: Don't prompt for confirmation to delete the role.

--8<-- "docs/cmd/_global.md"

## Remarks

For best performance use the `appObjectId` option to reference the Azure AD application registration from which to delete the role. If you use `appId` or `appName`, this command will first need to find the corresponding object ID for that application.

If the command finds multiple Azure AD application registrations with the specified app name, it will prompt you to disambiguate which app it should use, listing the discovered object IDs.

If the command finds multiple roles with the specified role name, it will prompt you to disambiguate which role it should use, listing the claim values.

If the role to be deleted is 'Enabled', this command will disable the role first and then delete.

## Examples

Delete role from a Azure AD application registration using object ID and role name options. Will prompt for confirmation before deleting the role.

```sh
m365 aad app role delete --appObjectId d75be2e1-0204-4f95-857d-51a37cf40be8 --name "Get Product"
```

Delete role from a Azure AD application registration using app (client) ID and role claim options. Will prompt for confirmation before deleting the role.

```sh
m365 aad app role delete --appId e75be2e1-0204-4f95-857d-51a37cf40be8 --claim "Product.Get"
```

Delete role from a Azure AD application registration using app name and role claim options. Will prompt for confirmation before deleting the role.

```sh
m365 aad app role delete --appName "My app" --claim "Product.Get"
```

Delete role from a Azure AD application registration using object ID and role id options. Will NOT prompt for confirmation before deleting the role.

```sh
m365 aad app role delete --appObjectId d75be2e1-0204-4f95-857d-51a37cf40be8 --id 15927ce6-1933-4b2f-b029-4dee3d53f4dd --confirm
```
