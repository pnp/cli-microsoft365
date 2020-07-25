# spo list get

Gets information about the specific list

## Usage

```sh
m365 spo list get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list to retrieve is located

`-i, --id [id]`
: ID of the list to retrieve information for. Specify either `id` or `title` but not both

`-t, --title [title]`
: Title of the list to retrieve information for. Specify either `id` or `title` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Return information about a list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list get --id 0cd891ef-afce-4e55-b836-fce03286cccf --webUrl https://contoso.sharepoint.com/sites/project-x
```

Return information about a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list get --title Documents --webUrl https://contoso.sharepoint.com/sites/project-x
```

## More information

- List REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint](https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint)