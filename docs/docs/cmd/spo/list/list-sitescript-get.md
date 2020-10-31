# spo list sitescript get

Extracts a site script from a SharePoint list

## Usage

```sh
m365 spo list sitescript get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list to extract the site script from is located

`-l, --listId [listId]`
: ID of the list to extract the site script from. Specify either `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list to extract the site script from. Specify either `listId` or `listTitle` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Extract a site script from an existing SharePoint list with title _ContosoList_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list sitescript get --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle ContosoList
```

Extract a site script from an existing SharePoint list with id _cc27a922-8224-4296-90a5-ebbc54da2e85_
located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list sitescript get --webUrl https://contoso.sharepoint.com/sites/project-x --listId cc27a922-8224-4296-90a5-ebbc54da2e85
```