# spo list contenttype remove

Removes content type from list

## Usage

```sh
m365 spo list contenttype remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`-l, --listId [listId]`
: ID of the list from which to remove the content type, specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list from which to remove the content type, specify `listId` or `listTitle` but not both

`-c, --contentTypeId <contentTypeId>`
: ID of the content type to remove from the list

`--confirm`
: Don't prompt for confirming removing the content type from the list

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Remove content type with ID _0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A_ from the list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list contenttype remove --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --contentTypeId 0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A
```

Remove content type with ID _0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A_ from the list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list contenttype remove --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --contentTypeId 0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A
```