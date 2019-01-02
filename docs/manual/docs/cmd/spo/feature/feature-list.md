# spo feature list

Lists features for site or site collection

## Usage

```sh
spo feature list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|Url of the site or site collection to retrieve the features from
`-s, --scope [scope]`|Scope of the features. Allowed values `Site|Web`. Default `Web`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To retrieve list of features, you have to first log in to a SharePoint Online site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

When using the text output type (default), the command lists only the values of the `DefinitionId` and `DisplayName` properties of the features. When setting the output type to JSON, all available properties are included in the command output.

## Examples

Return details about all user features located in site collection _https://contoso.sharepoint.com/sites/test_

```sh
spo feature list -u https://contoso.sharepoint.com/sites/test -s Site
```

Return details about all features located in site _https://contoso.sharepoint.com/sites/test_

```sh
spo feature list --url https://contoso.sharepoint.com/sites/test --scope Web
```

## More information

- Feature REST API resources: [https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj247054(v=office.15)#rest-resource-endpoint](https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj247054(v=office.15)#rest-resource-endpoint)