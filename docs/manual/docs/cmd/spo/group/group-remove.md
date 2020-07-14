# spo group remove

Removes site group from specific web

## Usage

```sh
spo group remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|Url of the web to remove the group from
`--id [id]`|Id of the site group to remove
`--name [name]`|Name of the site group to remove
`--confirm`|Confirm removal of user from site
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

Use either `id` or `name`, but not both

## Examples

Removes group with id _5_ for web _https://contoso.sharepoint.com/sites/mysite_

```sh
spo group remove --webUrl https://contoso.sharepoint.com/sites/mysite --id 5
```

Removes group with name _Team Site Owners_ for web _https://contoso.sharepoint.com/sites/mysite_

```sh
spo group remove --webUrl https://contoso.sharepoint.com/sites/mysite --name "Team Site Owners"
```