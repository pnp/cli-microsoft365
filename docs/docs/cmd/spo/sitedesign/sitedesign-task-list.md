# spo sitedesign task list

Lists site designs scheduled for execution on the specified site

## Usage

```sh
m365 spo sitedesign task list [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site for which to list site designs scheduled for execution

--8<-- "docs/cmd/_global.md"

## Examples

List site designs scheduled for execution on the specified site

```sh
m365 spo sitedesign task list --webUrl https://contoso.sharepoint.com/sites/team-a
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)
