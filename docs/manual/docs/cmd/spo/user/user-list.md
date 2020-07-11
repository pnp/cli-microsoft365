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
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets list of all users for web _https://contoso.sharepoint.com/sites/project-x_

```sh
spo user list --webUrl https://contoso.sharepoint.com/sites/project-x 
```
## More information

API Reference
```
GET https://contoso.sharepoint.com/sites/project-x/_api/web/SiteUsers
````
