# aad user get

Gets information about the specified user

## Usage

```sh
m365 aad user get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id [id]`
: The ID of the user to retrieve information for. Specify `id` or `userName` but not both

`-n, --userName [userName]`
: The name of the user to retrieve information for. Specify `id` or `userName` but not both

`-p, --properties [properties]`
: Comma-separated list of properties to retrieve

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

You can retrieve information about a user, either by specifying that user's id or user name (`userPrincipalName`), but not both.

If the user with the specified id or user name doesn't exist, you will get a `Resource 'xyz' does not exist or one of its queried reference-property objects are not present.` error.

## Examples

Get information about the user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad user get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844
```

Get information about the user with user name _AarifS@contoso.onmicrosoft.com_

```sh
m365 aad user get --userName AarifS@contoso.onmicrosoft.com
```

For the user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_ retrieve the user name, e-mail address and full name

```sh
m365 aad user get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --properties "userPrincipalName,mail,displayName"
```

## More information

- Microsoft Graph User properties: [https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/user#properties](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/user#properties)
