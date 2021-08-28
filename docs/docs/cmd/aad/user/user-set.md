# aad user set

Updates information about the specified user in Azure Active Directory AD

## Usage

```sh
m365 aad user set [options]
```

## Options

`-i, --objectId [objectId]`
: The object ID of the user to update. Specify `objectId` or `userPrincipalName` but not both

`-n, --userPrincipalName [userPrincipalName]`
: User principal name of the user to update. Specify `objectId` or `userPrincipalName` but not both

`--accountEnabled [accountEnabled]`
: Indicates whether the account is enabled

--8<-- "docs/cmd/_global.md"

## Remarks

You can retrieve information about a user, either by specifying that user's id or user name (`userPrincipalName`), but not both.

If the user with the specified id or user name doesn't exist, you will get a `Resource 'xyz' does not exist or one of its queried reference-property objects are not present.` error.

## Examples

Update specific property _department_ of user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad user set --objectId 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --Department IT
```

Update multiple properties of user with name _steve@contoso.onmicrosoft.com_

```sh
m365 aad user set --userPrincipalName steve@contoso.onmicrosoft.com --Department "Sales & Marketing" --CompanyName Contoso
```

Enable user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad user set --objectId 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --accountEnabled true
```

Disable user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad user set --objectId 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --accountEnabled false
```

Enable user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad user set --objectId 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --accountEnabled true
```
