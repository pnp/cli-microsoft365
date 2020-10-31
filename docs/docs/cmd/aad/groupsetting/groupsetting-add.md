# aad groupsetting add

Creates a group setting

## Usage

```sh
m365 aad groupsetting add [options]
```

## Options

`-h, --help`
: output usage information

`-i, --templateId <templateId>`
: The ID of the group setting template to use to create the group setting

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

To create a group setting, you have to specify the ID of the group setting template that should be used to create the setting. You can retrieve the ID of the template using the [aad groupsettingtemplate list](../groupsettingtemplate/groupsettingtemplate-list.md) command.

To specify values for the different properties specified in the group setting template, include additional options that match the property in the group setting template. For example `--ClassificationList 'HBI, MBI, LBI, GDPR'` will set the list of classifications to use on modern SharePoint sites.

Each group setting template specifies default value for each property. If you don't specify a value for the particular property yourself, the default value from the group setting template will be used. To find out which properties are available for the particular group setting template, use the [aad groupsettingtemplate get](../groupsettingtemplate/groupsettingtemplate-get.md) command.

If the specified templateId doesn't reference a valid group setting template, you will get a _Resource 'xyz' does not exist or one of its queried reference-property objects are not present._ error.

If you try to add a group setting using a template, for which a setting already exists, you will get a _A conflicting object with one or more of the specified property values is present in the directory._ error.

## Examples

Configure classification for modern SharePoint sites

```sh
m365 aad groupsetting add --templateId 62375ab9-6b52-47ed-826b-58e47e0e304b --UsageGuidelinesUrl https://contoso.sharepoint.com/sites/compliance --ClassificationList 'HBI, MBI, LBI, GDPR' --DefaultClassification MBI
```