# spo contenttype remove

Removes a content type from a site if not in use

## Usage

```sh
spo contenttype remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|Absolute URL of the site where the content type is located
`-i, --contentTypeId <contentTypeId>`|The ID of the content type to remove
`-n, --name [contentTypeName]`|Content type name to remove if the content type id is not known. Either id or name must be specified but not both paramters
`--confirm`|Don't prompt for confirming removal of the content type
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To remove a column from a content type, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

Content Types in use may not be deleted and will return an error. SharePoint will not allow a content type to be removed unless any dependent objects are also emptied from the recycle bin incluing the second-stage recycle bin.

## Examples

Remove content type with ContentTypeID _0x01007926A45D687BA842B947286090B8F67D_ from a site with URL _https://contoso.sharepoint.com_

```sh
spo contenttype remove --id "0x01007926A45D687BA842B947286090B8F67D" --webUrl https://contoso.sharepoint.com
```

Remove content type with the name _My Content Type_ from a site with URL _https://contoso.sharepoint.com_

```sh
spo contenttype remove --name "My Content Type" --webUrl https://contoso.sharepoint.com --confirm
```

Remove content type with the name _My Content Type_ from a site with URL _https://contoso.sharepoint.com_ without prompting for confirmation

```sh
spo contenttype remove --name "My Content Type" --webUrl https://contoso.sharepoint.com --confirm
```