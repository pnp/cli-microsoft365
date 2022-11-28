# spo page control set

Updates web part data or properties of a control on a modern page

## Usage

```sh
m365 spo page control set [options]
```

## Options

`-i, --id <id>`
: ID of the control to update properties of.

`-n, --pageName <pageName>`
: Name of the page where the control is located.

`-u, --webUrl <webUrl>`
: URL of the site where the page is located.

`--webPartData [webPartData]`
: JSON string with web part data as retrieved from the web part maintenance mode. Specify either `webPartProperties` or `webPartData` but not both.

`--webPartProperties [webPartProperties]`
: JSON string with web part data as retrieved from the web part maintenance mode. Specify either `webPartProperties` or `webPartData` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `pageName` doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

!!! warning "Escaping JSON in PowerShell"
    When using the `--webPartProperties` option it's possible to enter a JSON string. In PowerShell 5 to 7.2 [specific escaping rules](./../../../user-guide/using-cli.md#escaping-double-quotes-in-powershell) apply due to an issue. Remember that you can also use [file tokens](./../../../user-guide/using-cli.md#passing-complex-content-into-cli-options) instead.

## Examples

Update web part data for the control, placed on a modern page

```sh
m365 spo page control set --id 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1 --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --webPartData '{"title":"New WP Title","properties": {"description": "New description"}}'
```

Update web part properties for the control, placed on a modern page

```sh
m365 spo page control set --id 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1 --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --webPartProperties '{"description": "New description"}'
```

## Response

The command won't return a response on success.
