# spo list label set

Sets classification label on the specified list

## Usage

```sh
spo list label set  [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the list is located
`--label <label>`|The label to set on the list
`-t, --listTitle [listTitle]`|Title of the list where the field is located. Specify only one of `listTitle`, `listId` or `listUrl`
`-l, --listId [listId]`|ID of the list where the field is located. Specify only one of `listTitle`, `listId` or `listUrl`
`--listUrl [listUrl]`|Server- or web-relative URL of the list where the field is located. Specify only one of `listTitle`, `listId` or `listUrl`
`--syncToItems [syncToItems]`|Specify, to set the label on all items in the list
`--blockDelete [blockDelete]`|Specify, to disallow deleting items in the list
`--blockEdit [blockEdit]`|Specify, to disallow editing items in the list
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To set a list classification label, you have to first log in to SharePoint using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Sets classification label "Confidential" for list _Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_
      

```sh
spo list label set --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'Shared Documents' --label 'Confidential'
```

Sets classification label "Confidential" and applies 'blockEdit','blockDelete' and 'syncToItems' flags for list _Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_
      

```sh
spo list label set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'Documents' --label 'Confidential' --blockEdit --blockDelete --syncToItems
```

## More information:

PnP PowerShell alternative: [https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/set-pnplabel?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/set-pnplabel?view=sharepoint-ps)