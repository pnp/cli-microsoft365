# spo sitedesign rights grant

Grants access to a site design for one or more principals

## Usage

```sh
m365 spo sitedesign rights grant [options]
```

## Options

`-i, --siteDesignId <siteDesignId>`
: The ID of the site design to grant rights on

`-p, --principals <principals>`
: Comma-separated list of principals to grant view rights. Principals can be users or mail-enabled security groups in the form of `alias` or `alias@<domain name>.com`

`-r, --rights <rights>`
: Rights to grant to principals. Available values `View`

--8<-- "docs/cmd/_global.md"

## Examples

Grant user with alias _PattiF_ view permission to the specified site design

```sh
m365 spo sitedesign rights grant --siteDesignId 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --principals PattiF --rights View
```

Grant users with aliases _PattiF_ and _AdeleV_ view permission to the specified site design

```sh
m365 spo sitedesign rights grant --siteDesignId 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --principals "PattiF,AdeleV" --rights View
```

Grant user with email _PattiF@contoso.com_ view permission to the specified site design

```sh
m365 spo sitedesign rights grant --siteDesignId 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --principals PattiF@contoso.com --rights View
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)
