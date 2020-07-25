# spo folder rename

Renames a folder

## Usage

```sh
m365 spo folder rename [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site where the folder to be renamed is located

`-f, --folderUrl <folderUrl>`
: Site-relative URL of the folder (including the folder)

`-n, --name`
: New name for the target folder

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Renames a folder with site-relative URL _/Shared Documents/My Folder 1_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo folder rename --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl '/Shared Documents/My Folder 1' --name 'My Folder 2'
```