# spo user list

Lists all the users within specific web

## Usage

```sh
spo user list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the web to list the users from
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Get list of users in web _https://contoso.sharepoint.com/sites/project-x_

```sh
spo user list --webUrl https://contoso.sharepoint.com/sites/project-x
```
