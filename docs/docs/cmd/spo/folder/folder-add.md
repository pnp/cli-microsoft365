# spo folder add

Creates a folder within a parent folder

## Usage

```sh
m365 spo folder add [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site where the folder will be created

`-p, --parentFolderUrl <parentFolderUrl>`
: Site-relative URL of the parent folder

`-n, --name <name>`
: Name of the new folder to be created

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Creates folder in a parent folder with site relative url _/Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo folder add --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/Shared Documents' --name 'My Folder Name'
```