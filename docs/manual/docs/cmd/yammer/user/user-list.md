# yammer user list

Returns users from the current network

## Usage

```sh
yammer user list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-g, --groupId [groupId]`|Returns users within a given group
`-l, --letter [letter]`|Returns users with usernames beginning with the given character
`--reverse`|Returns users in reverse sorting order
`--limit [limit]`|Limits the users returned
`--sortBy [sortBy]`|Returns users sorted by a number of messages or followers, instead of the default behavior of sorting alphabetically. Allowed values are `messages,followers`
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the Office 365 CLI the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples
  
Returns all Yammer network users

```sh
yammer user list
```

Returns all Yammer network users with usernames beginning with "a"

```sh
yammer user list --letter a
```

Returns all Yammer network users sorted alphabetically in descending order

```sh
yammer user list --reverse
```

Returns the first 10 Yammer network users within the group 5785177

```sh
user list --groupId 5785177 --limit 10
```
