# spo web set

Updates subsite properties

## Usage

```sh
spo web set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the subsite to update
`-t, --title [title]`|New title for the subsite
`-d, --description [description]`|New description for the subsite
`--siteLogoUrl [siteLogoUrl]`|New site logo URL for the subsite
`--quickLaunchEnabled [quickLaunchEnabled]`|Set to `true` to enable quick launch and to `false` to disable it
`--headerLayout [headerLayout]`|Configures the site header. Allowed values `standard|compact`
`--headerEmphasis [headerEmphasis]`|Configures the site header background. Allowed values `0|1|2|3`
`--megaMenuEnabled [megaMenuEnabled]`|Set to `true` to change the menu style to megamenu. Set to `false` to use the cascading menu style
`--footerEnabled [footerEnabled]`|Set to `true` to enable footer and to `false` to disable it
`--searchScope [searchScope]`|Search scope to set in the site. Allowed values `DefaultScope|Tenant|Hub|Site`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Update subsite title

```sh
spo web set --webUrl https://contoso.sharepoint.com/sites/team-a --title Team-a
```

Hide quick launch on the subsite

```sh
spo web set --webUrl https://contoso.sharepoint.com/sites/team-a --quickLaunchEnabled false
```

Set site header layout to compact

```sh
spo web set --webUrl https://contoso.sharepoint.com/sites/team-a --headerLayout compact
```

Set site header color to primary theme background color

```sh
spo web set --webUrl https://contoso.sharepoint.com/sites/team-a --headerEmphasis 0
```

Enable megamenu in the site

```sh
spo web set --webUrl https://contoso.sharepoint.com/sites/team-a --megaMenuEnabled true
```

Hide footer in the site

```sh
spo web set --webUrl https://contoso.sharepoint.com/sites/team-a --footerEnabled false
```

Set search scope to tenant scope

```sh
spo web set --webUrl https://contoso.sharepoint.com/sites/team-a --searchScope tenant
```