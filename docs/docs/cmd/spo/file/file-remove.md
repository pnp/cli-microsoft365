# spo file remove

Removes the specified file

## Usage

```sh
m365 spo file remove [options]
```

## Options

`-h, --help`
: output usage information

`-w, --webUrl <webUrl>`
: URL of the site where the file to remove is located

`-i, --id [id]`
: The ID of the file to remove. Specify either `id` or `url` but not both

`-u, --url [url]`
: The server- or site-relative URL of the file to remove. Specify either `id` or `url` but not both

`--recycle`
: Recycle the file instead of actually deleting it

`--confirm`
: Don't prompt for confirming removing the file

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Remove the file with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file remove --webUrl https://contoso.sharepoint.com/sites/project-x --id 0cd891ef-afce-4e55-b836-fce03286cccf
```

Remove the file with site-relative URL _SharedDocuments/Test.docx_ from located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file remove --webUrl https://contoso.sharepoint.com/sites/project-x --url SharedDocuments/Test.docx
```

Move the file with server-relative URL _/sites/project-x/SharedDocuments/Test.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_ to the recycle bin

```sh
m365 spo file remove --webUrl https://contoso.sharepoint.com/sites/project-x --url /sites/project-x/SharedDocuments/Test.docx --recycle
```