# spo page control set

Updates web part data or properties of a control on a modern page

## Usage

```sh
m365 spo page control set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: ID of the control to update properties of

`-n, --name <name>`
: Name of the page where the control is located

`-u, --webUrl <webUrl>`
: URL of the site where the page is located

`--webPartData [webPartData]`
: JSON string with web part data as retrieved from the web part maintenance mode. Specify `webPartProperties` or `webPartData` but not both

`--webPartProperties [webPartProperties]`
: JSON string with web part data as retrieved from the web part maintenance mode. Specify `webPartProperties` or `webPartData` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified `name` doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

When specifying the JSON string with web part properties on Windows, you have to escape double quotes in a specific way. Considering the following value for the _webPartProperties_ option: `{"Foo":"Bar"}`, you should specify the value as \`"{""Foo"":""Bar""}"\`. In addition, when using PowerShell, you should use the `--%` argument.

## Examples

Update web part data for the control with ID _3ede60d3-dc2c-438b-b5bf-cc40bb2351e1_ placed on a modern page with name _home.aspx_

```sh
m365 spo page control set --id 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1 --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx --webPartData '{"title":"New WP Title","properties": {"description": "New description"}}'
```

Update web part properties for the control with ID _3ede60d3-dc2c-438b-b5bf-cc40bb2351e1_ placed on a modern page with name _home.aspx_

```sh
m365 spo page control set --id 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1 --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx --webPartProperties '{"description": "New description"}'
```
