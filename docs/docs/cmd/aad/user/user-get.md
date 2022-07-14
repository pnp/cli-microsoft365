# aad user get

Gets information about the specified user

## Usage

```sh
m365 aad user get [options]
```

## Options

`-i, --id [id]`
: The ID of the user to retrieve information for. Use either `id`, `userName` or `email`, but not all.

`-n, --userName [userName]`
: The name of the user to retrieve information for. Use either `id`, `userName` or `email`, but not all.

`--email [email]`
: The email of the user to retrieve information for. Use either `id`, `userName` or `email`, but not all.

`-p, --properties [properties]`
: Comma-separated list of properties to retrieve

--8<-- "docs/cmd/_global.md"

## Remarks

You can retrieve information about a user, either by specifying that user's id, user name (`userPrincipalName`), or email (`mail`), but not all.

If the user with the specified id, user name, or email doesn't exist, you will get a `Resource 'xyz' does not exist or one of its queried reference-property objects are not present.` error.

## Examples

Get information about the user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad user get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844
```

Get information about the user with user name _AarifS@contoso.onmicrosoft.com_

```sh
m365 aad user get --userName AarifS@contoso.onmicrosoft.com
```

Get information about the user with email _AarifS@contoso.onmicrosoft.com_

```sh
m365 aad user get --email AarifS@contoso.onmicrosoft.com
```

For the user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_ retrieve the user name, e-mail address and full name

```sh
m365 aad user get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --properties "userPrincipalName,mail,displayName"
```

Get information about the currently logged user using the Id token

```sh
m365 aad user get --id "@meId"
```

Get information about the currently logged in user using the UserName token

```sh
m365 aad user get --userName "@meUserName"
```

## More information

- Microsoft Graph User properties: [https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/user#properties](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/user#properties)
