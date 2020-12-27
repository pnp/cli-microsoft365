# spo sitedesign run list

Lists information about site designs applied to the specified site

## Usage

```sh
m365 spo sitedesign run list [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site for which to list applied site designs

`-i, --siteDesignId [siteDesignId]`
: The ID of the site design for which to display information

--8<-- "docs/cmd/_global.md"

## Examples

List site designs applied to the specified site

```sh
m365 spo sitedesign run list --webUrl https://contoso.sharepoint.com/sites/team-a
```

List information about the specified site design applied to the specified site

```sh
m365 spo sitedesign run list --webUrl https://contoso.sharepoint.com/sites/team-a --siteDesignId 6ec3ca5b-d04b-4381-b169-61378556d76e
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)