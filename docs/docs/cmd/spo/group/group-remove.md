# spo group remove

Removes group from specific web

## Usage

```sh
m365 spo group remove [options]
```

## Options

Option|Description
------|-----------
`-h, --help`|output usage information
`-u, --webUrl <webUrl>`|Url of the web to remove the group from
`--id [id]`|ID of the group to remove. Use ID or name but not both
`--name [name]`|Name of the group to remove. Use ID or name but not both
`--confirm`|Confirm removal of the group
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Removes group with id _5_ from web _https://contoso.sharepoint.com/sites/mysite_

```sh
m365 spo group remove --webUrl https://contoso.sharepoint.com/sites/mysite --id 5
```

Removes group with name _Team Site Owners_ from web _https://contoso.sharepoint.com/sites/mysite_

```sh
m365 spo group remove --webUrl https://contoso.sharepoint.com/sites/mysite --name "Team Site Owners"
```
