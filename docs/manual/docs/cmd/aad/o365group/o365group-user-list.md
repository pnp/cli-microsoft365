# aad o365group user list

Lists users for the specified Office 365 group or Microsoft Teams team

## Usage

```sh
aad o365group user list [options]
```

## Alias

```sh
teams user list
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --groupId [groupId]`|The ID of the Office 365 group for which to list users
`--teamId [teamId]`|The ID of the Microsoft Teams team for which to list users
`-r, --role [role]`|Filter the results to only users with the given role: `Owner,Member,Guest`
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

List all users and their role in the specified Office 365 group

```sh
aad o365group user list --groupId '00000000-0000-0000-0000-000000000000'
```

List all owners and their role in the specified Office 365 group

```sh
aad o365group user list --groupId '00000000-0000-0000-0000-000000000000' --role Owner
```

 List all guests and their role in the specified Office 365 group

```sh
aad o365group user list --groupId '00000000-0000-0000-0000-000000000000' --role Guest
```

List all users and their role in the specified Microsoft teams team

```sh
teams user list --teamId '00000000-0000-0000-0000-000000000000'
```

List all owners and their role in the specified Microsoft teams team

```sh
teams user list --teamId '00000000-0000-0000-0000-000000000000' --role Owner
```