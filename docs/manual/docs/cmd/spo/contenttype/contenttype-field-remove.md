# spo contenttype field remove

Removes field link from a site or list content type.

## Usage

```sh
spo contenttype field remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|Absolute URL of the site where the content type is located
`-i, --contentTypeId <contentTypeId>`|ID of the content type on which the field link should be removed
`-f, --fieldId <fieldId>`|ID of the field link which should be remove from the content type
`-c, --updateChildContentTypes <true|false>`|Specifies if the child content types should be updated. Option is ignored if specified list title
`-l, --listTitle <listTitle>`|List title where the content type exists
`--confirm`|Don't prompt for confirming removal of a field link from content type
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To remove a field link from a content type, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.


## Examples

Remove field link from a site content type without prompt

```sh

spo contenttype field remove -webUrl https://contoso.sharepoint.com/sites/portal --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A6" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --confirm
```

Remove field link from a site content type with the child content types update and with prompt

```sh
spo contenttype field remove -webUrl https://contoso.sharepoint.com/sites/portal --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A6" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --updateChildContentTypes true
```

Remove field link from a list content type with prompt

```sh
spo contenttype field remove -webUrl https://contoso.sharepoint.com/sites/portal --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A60062F089A38C867747942DB2C3FC50FF6A" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --listTitle ListName
```