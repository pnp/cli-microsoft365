# accesstoken get

Gets access token for the specified resource

## Usage

```sh
accesstoken get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-r, --resource <resource>`|The resource for which to retrieve an access token
`--new`|Retrieve a new access token to ensure that it's valid for as long as possible
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

The `accesstoken get` command returns an access token for the specified resource. If an access token has been previously retrieved and is still valid, the command will return the cached token. If you want to ensure that the returned access token is valid for as long as possible, you can force the command to retrieve a new access token by using the `--new` option.

## Examples

Get access token for the Microsoft Graph

```sh
accesstoken get --resource https://graph.microsoft.com
```

Get a new access token for SharePoint Online

```sh
accesstoken get --resource https://contoso.sharepoint.com --new
```
