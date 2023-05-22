# spo applicationcustomizer set

Updates an existing application customizer on a site

## Usage

```sh
m365 spo applicationcustomizer set [options]
```

## Options

`-u, --webUrl <webUrl>`
: The site to update the Application Customizer on.

`-t, --title [title]`
: The title of the Application Customizer to update. Specify either `title`, `id` or `clientSideComponentId`.

`-i, --id [id]`
: The id of the Application Customizer to update. Specify either `title`, `id` or `clientSideComponentId`.

`-c, --clientSideComponentId [clientSideComponentId]`
: The Client Side Component Id (GUID) of the Application Customizer to update. Specify either `title`, `id` or `clientSideComponentId`.

`--newTitle [newTitle]`
: The new title of the Application Customizer.

`--clientSideComponentProperties  [clientSideComponentProperties]`
: The Client Side Component properties of the Application Customizer.

`-s, --scope [scope]`
: The scope where to lookup the Application Customizer. Allowed values: `Site`, `Web`, and `All`. Defaults to `All`.

--8<-- "docs/cmd/_global.md"

## Remarks

Running this command from the Windows Command Shell (cmd.exe) or PowerShell for Windows OS XP, 7, 8, 8.1 without bash installed might require additional formatting for clientSideComponentProperties option that has JSON value because the command shell treat quotes differently. For example, this is how Application Customizer can be updated from the Windows cmd.exe:

```sh
m365 spo applicationcustomizer set --webUrl https://contoso.sharepoint.com/sites/sales --id b41916e7-e69d-467f-b37f-ff8ecf8f99f2 --newTitle "Some customizer" --clientSideComponentProperties '{\"testMessage\":\"Test message\"}'
```

Note, how the clientSideComponentProperties option (--clientSideComponentProperties) has escaped double quotes `'{\"testMessage\":\"Test message\"}'` compared to execution from bash `'{"testMessage":"Test message"}'`.

!!! warning "Escaping JSON in PowerShell"
    When using the `--clientSideComponentProperties` option it's possible to enter a JSON string. In PowerShell 5 to 7.2 [specific escaping rules](./../../../user-guide/using-cli.md#escaping-double-quotes-in-powershell) apply due to an issue. Remember that you can also use [file tokens](./../../../user-guide/using-cli.md#passing-complex-content-into-cli-options) instead.

## Examples

Updates the title of an application customizer on the sales site

```sh
m365 spo applicationcustomizer set --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --newTitle "Some customizer" --webUrl https://contoso.sharepoint.com/sites/sales
```

Updates the properties of an application customizer on the sales site

```sh
m365 spo applicationcustomizer set --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --clientSideComponentProperties '{ "testMessage": "Test message" }' --webUrl https://contoso.sharepoint.com/sites/sales
```

## Response

The command won't return a response on success.
