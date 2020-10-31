# spo customaction add

Adds a user custom action for site or site collection

## Usage

```sh
m365 spo customaction add [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: Url of the site or site collection to add the custom action

`-n, --name <name>`
: The name of the custom action

`-t, --title <title>`
: The title of the custom action

`-l, --location <location>`
: The actual location where this custom action need to be added like `CommandUI.Ribbon`

`-g, --group [group]`
: The group where this custom action needs to be added like `SiteActions`

`-d, --description [description]`
: The description of the custom action

`--sequence [sequence]`
: Sequence of this CustomAction being injected. Use when you have a specific sequence with which to have multiple CustomActions being added to the page

`--actionUrl [actionUrl]`
: The URL, URI or JavaScript function associated with the action. URL example `~site/_layouts/sampleurl.aspx` or `~sitecollection/_layouts/sampleurl.aspx`

`--imageUrl [imageUrl]`
: The URL of the image associated with the custom action

`-e, --commandUIExtension [commandUIExtension]`
: XML fragment that determines user interface properties of the custom action

`--registrationId [registrationId]`
: Specifies the identifier of the list or item content type that this action is associated with, or the file type or programmatic identifier

`--registrationType [registrationType]`
: Specifies the type of object associated with the custom action. Allowed values `None,List,ContentType,ProgId,FileType`. Default `None`

`--rights [rights]`
: A case sensitive string array that contain the permissions needed for the custom action. Allowed values `EmptyMask,ViewListItems,AddListItems,EditListItems,DeleteListItems,ApproveItems,OpenItems,ViewVersions,DeleteVersions,CancelCheckout,ManagePersonalViews,ManageLists,ViewFormPages,AnonymousSearchAccessList,Open,ViewPages,AddAndCustomizePages,ApplyThemeAndBorder,ApplyStyleSheets,ViewUsageData,CreateSSCSite,ManageSubwebs,CreateGroups,ManagePermissions,BrowseDirectories,BrowseUserInfo,AddDelPrivateWebParts,UpdatePersonalWebParts,ManageWeb,AnonymousSearchAccessWebLists,UseClientIntegration,UseRemoteAPIs,ManageAlerts,CreateAlerts,EditMyUserInfo,EnumeratePermissions,FullMask`. Default `EmptyMask`

`-s, --scope [scope]`
: Scope of the custom action. Allowed values `Site,Web`. Default `Web`

`--scriptBlock [scriptBlock]`
: Specifies a block of script to be executed. This attribute is only applicable when the Location attribute is set to ScriptLink

`--scriptSrc [scriptSrc]`
: Specifies a file that contains script to be executed. This attribute is only applicable when the Location attribute is set to ScriptLink

`-c, --clientSideComponentId [clientSideComponentId]`
: The Client Side Component Id (GUID) of the custom action

`-p, --clientSideComponentProperties [clientSideComponentProperties]`
: The Client Side Component Properties of the custom action. Specify values as a JSON string : `'{"testMessage":"Test message"}'`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Running this command from the Windows Command Shell (cmd.exe) or PowerShell for Windows OS XP, 7, 8, 8.1 without bash installed might require additional formatting for command options that have JSON, XML or JavaScript values because the command shell treat quotes differently. For example, this is how ApplicationCustomizer user custom action can be created from the Windows cmd.exe:

```sh
m365 spo customaction add -u https://contoso.sharepoint.com/sites/test -t "YourAppCustomizer" -n "YourName" -l "ClientSideExtension.ApplicationCustomizer" -c b41916e7-e69d-467f-b37f-ff8ecf8f99f2 -p '{\"testMessage\":\"Test message\"}'
```

Note, how the clientSideComponentProperties option (-p) has escaped double quotes `'{\"testMessage\":\"Test message\"}'` compared to execution from bash `'{"testMessage":"Test message"}'`.

The `--rights` option accepts **case sensitive** values.

## Examples

Adds tenant-wide SharePoint Framework Application Customizer extension in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction add -u https://contoso.sharepoint.com/sites/test -t "YourAppCustomizer" -n "YourName" -l "ClientSideExtension.ApplicationCustomizer" -c b41916e7-e69d-467f-b37f-ff8ecf8f99f2 -p '{"testMessage":"Test message"}'
```

Adds tenant-wide SharePoint Framework **modern List View** Command Set extension in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction add -u https://contoso.sharepoint.com/sites/test -t "YourCommandSet" -n "YourName" -l "ClientSideExtension.ListViewCommandSet" -c db3e6e35-363c-42b9-a254-ca661e437848 -p '{"sampleTextOne":"One item is selected in the list.", "sampleTextTwo":"This command is always visible."}' --registrationId 100 --registrationType List
```

Creates url custom action in the SiteActions menu in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction add -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "Microsoft.SharePoint.StandardMenu" -g "SiteActions" --actionUrl "~site/SitePages/Home.aspx" --sequence 100
```

Creates custom action in **classic** Document Library edit context menu in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction add -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "EditControlBlock" --actionUrl "javascript:(function(){ return console.log('CLI for Microsoft 365 rocks!'); })();" --registrationId 101 --registrationType List
```

Creates ScriptLink custom action with script source in **classic pages** in site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction add -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "ScriptLink" --scriptSrc "~sitecollection/SiteAssets/YourScript.js" --sequence 101 -s Site
```

Creates ScriptLink custom action with script block in **classic pages** in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction add -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "ScriptLink" --scriptBlock "(function(){ return console.log('Hello CLI for Microsoft 365!'); })();" --sequence 102
```

Creates **classic List View** custom action located in the Ribbon in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction add -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "CommandUI.Ribbon" --commandUIExtension '<CommandUIExtension><CommandUIDefinitions><CommandUIDefinition Location="Ribbon.List.Share.Controls._children"><Button Id="Ribbon.List.Share.GetItemsCountButton" Alt="Get list items count" Sequence="11" Command="Invoke_GetItemsCountButtonRequest" LabelText="Get Items Count" TemplateAlias="o1" Image32by32="_layouts/15/images/placeholder32x32.png" Image16by16="_layouts/15/images/placeholder16x16.png" /></CommandUIDefinition></CommandUIDefinitions><CommandUIHandlers><CommandUIHandler Command="Invoke_GetItemsCountButtonRequest" CommandAction="javascript: alert(ctx.TotalListItems);" EnabledScript="javascript: function checkEnable() { return (true);} checkEnable();"/></CommandUIHandlers></CommandUIExtension>'
```

Creates custom action with delegated rights in the SiteActions menu in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction add -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "Microsoft.SharePoint.StandardMenu" -g "SiteActions" --actionUrl "~site/SitePages/Home.aspx" --rights "AddListItems,DeleteListItems,ManageLists"
```

## More information

- UserCustomAction REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction](https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction)
- UserCustomAction Locations and Group IDs: [https://msdn.microsoft.com/en-us/library/office/bb802730.aspx](https://msdn.microsoft.com/en-us/library/office/bb802730.aspx)
- UserCustomAction Element: [https://msdn.microsoft.com/en-us/library/office/ms460194.aspx](https://msdn.microsoft.com/en-us/library/office/ms460194.aspx)
- UserCustomAction Rights: [https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.spbasepermissions.aspx](https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.spbasepermissions.aspx)