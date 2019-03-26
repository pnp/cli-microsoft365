# spo sitedesign task list

Lists site designs scheduled for execution on the specified site

## Usage

```sh
spo sitedesign task list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site for which to list site designs scheduled for execution
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

List site designs scheduled for execution on the specified site

```sh
spo sitedesign task list --webUrl https://contoso.sharepoint.com/sites/team-a
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)