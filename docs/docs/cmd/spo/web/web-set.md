# spo web set

Updates subsite properties

## Usage

```sh
m365 spo web set [options]
```

## Options

`-u, --url <url>`
: URL of the subsite to update

`-t, --title [title]`
: New title for the subsite

`-d, --description [description]`
: New description for the subsite

`--siteLogoUrl [siteLogoUrl]`
: New site logo URL for the subsite. Set to empty string to reset to default.

`--quickLaunchEnabled [quickLaunchEnabled]`
: Set to `true` to enable quick launch and to `false` to disable it

`--headerLayout [headerLayout]`
: Configures the site header. Allowed values `standard,compact`

`--headerEmphasis [headerEmphasis]`
: Configures the site header background. Allowed values `0,1,2,3`

`--megaMenuEnabled [megaMenuEnabled]`
: Set to `true` to change the menu style to megamenu. Set to `false` to use the cascading menu style

`--footerEnabled [footerEnabled]`
: Set to `true` to enable footer and to `false` to disable it

`--searchScope [searchScope]`
: Search scope to set in the site. Allowed values `DefaultScope,Tenant,Hub,Site`

`--welcomePage [welcomePage]`
: Site-relative URL of the welcome page for the site

--8<-- "docs/cmd/_global.md"

## Remarks

Next to updating web properties corresponding to the options of this command, you can update the value of any other web property using its CSOM name, eg. `--AllowAutomaticASPXPageIndexing`. At this moment, the CLI supports properties of types `Boolean`, `String` and `Int32`.

## Examples

Update subsite title

```sh
m365 spo web set --url https://contoso.sharepoint.com/sites/team-a --title Team-a
```

Hide quick launch on the subsite

```sh
m365 spo web set --url https://contoso.sharepoint.com/sites/team-a --quickLaunchEnabled false
```

Set site header layout to compact

```sh
m365 spo web set --url https://contoso.sharepoint.com/sites/team-a --headerLayout compact
```

Set site header color to primary theme background color

```sh
m365 spo web set --url https://contoso.sharepoint.com/sites/team-a --headerEmphasis 0
```

Enable megamenu in the site

```sh
m365 spo web set --url https://contoso.sharepoint.com/sites/team-a --megaMenuEnabled true
```

Hide footer in the site

```sh
m365 spo web set --url https://contoso.sharepoint.com/sites/team-a --footerEnabled false
```

Set search scope to tenant scope

```sh
m365 spo web set --url https://contoso.sharepoint.com/sites/team-a --searchScope tenant
```

Set welcome page for the web

```sh
m365 spo web set  --url https://contoso.sharepoint.com/sites/team-a --welcomePage "SitePages/new-home.aspx"
```

## Response

The command won't return a response on success.

## More information

- Web properties: [https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee545886(v=office.15)](https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee545886(v=office.15))
