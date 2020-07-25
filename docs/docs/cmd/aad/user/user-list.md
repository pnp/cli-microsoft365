# aad user list

Lists users matching specified criteria

## Usage

```sh
m365 aad user list [options]
```

## Options

`-h, --help`
: output usage information

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

Using the `--properties` option, you can specify a comma-separated list of user properties to retrieve from the Microsoft Graph. If you don't specify any properties, the command will retrieve user's display name and account name.

To filter the list of users, include additional options that match the user property that you want to filter with. For example `--displayName Patt` will return all users whose `displayName` starts with `Patt`. Multiple filters will be combined using the `and` operator.

## Examples

List all users in the tenant

```sh
m365 aad user list
```

List all users in the tenant. For each one return the display name and e-mail address

```sh
m365 aad user list --properties "displayName,mail"
```

Show users whose display name starts with _Patt_

```sh
m365 aad user list --displayName Patt
```

Show all account managers whose display name starts with _Patt_

```sh
m365 aad user list --displayName Patt --jobTitle 'Account manager'
```

## More information

- Microsoft Graph User properties: [https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/user#properties](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/user#properties)
