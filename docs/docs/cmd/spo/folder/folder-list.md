# spo folder list

Returns all folders under the specified parent folder

## Usage

```sh
spo folder list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the folders to list are located
`-p, --parentFolderUrl <parentFolderUrl>`|Site-relative URL of the parent folder
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets list of folders under a parent folder with site-relative url _/Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo folder list --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/Shared Documents'
```