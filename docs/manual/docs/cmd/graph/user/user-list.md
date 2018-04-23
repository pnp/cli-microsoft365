# graph user list

Lists users matching specified criteria

## Usage

```sh
graph user list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --properties [properties]`|Comma-separated list of properties to retrieve
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

To list users matching specific criteria, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

Using the `--properties` option, you can specify a comma-separated list of user properties to retrieve from the Microsoft Graph. If you don't specify any properties, the command will retrieve user's display name and account name.

To filter the list of users, include additional options that match the user property that you want to filter with. For example `--displayName Patt` will return all users whose `displayName` starts with `Patt`. Multiple filters will be combined using the `and` operator.

## Examples

List all users in the tenant

```sh
graph user list
```

List all users in the tenant. For each one return the display name and e-mail address

```sh
graph user list --properties displayName,mail
```

Show users whose display name starts with _Patt_

```sh
graph user list --displayName Patt
```

Show all account managers whose display name starts with _Patt_

```sh
graph user list --displayName Patt --jobTitle 'Account manager'
```

## More information

- Microsoft Graph User properties: [https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/user#properties](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/user#properties)