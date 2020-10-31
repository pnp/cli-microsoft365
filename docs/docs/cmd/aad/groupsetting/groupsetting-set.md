# aad groupsetting set

Updates the particular group setting

## Usage

```sh
m365 aad groupsetting set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The ID of the group setting to update

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

To update a group setting, you have to specify the ID of the group setting. You can retrieve the ID of the group setting using the [aad groupsetting list](./groupsetting-list.md) command.

To update values for the different properties specified in the group setting, include additional options that match the property in the group setting. For example `--ClassificationList 'HBI, MBI, LBI, GDPR'` will set the list of classifications to use on modern SharePoint sites.

If you don't specify a value for the particular property, it will remain unchanged. To find out which properties are available for the particular group setting, use the [aad groupsetting get](./groupsetting-get.md) command.

If the specified id doesn't reference a valid group setting, you will get a _Resource 'xyz' does not exist or one of its queried reference-property objects are not present._ error.

## Examples

Configure classification for modern SharePoint sites

```sh
m365 aad groupsetting set --id c391b57d-5783-4c53-9236-cefb5c6ef323 --UsageGuidelinesUrl https://contoso.sharepoint.com/sites/compliance --ClassificationList 'HBI, MBI, LBI, GDPR' --DefaultClassification MBI
```