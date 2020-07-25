# spo list list

Gets all lists within the specified site

## Usage

```sh
m365 spo list list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the lists to retrieve are located

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Return all lists located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list list --webUrl https://contoso.sharepoint.com/sites/project-x
```

## More information

- List REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint](https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint)