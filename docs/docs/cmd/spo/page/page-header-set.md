# spo page header set

Sets modern page header

## Usage

```sh
m365 spo page header set [options]
```

## Options

`-h, --help`
: output usage information

`-n, --pageName <pageName>`
: Name of the page to set the header for

`-u, --webUrl <webUrl>`
: URL of the site where the page to update is located

`-t, --type [type]`
: Type of header, allowed values `None,Default,Custom`. Default `Default`

`--imageUrl [imageUrl]`
: Server-relative URL of the image to use in the header. Image must be stored in the same site collection as the page

`--altText [altText]`
: Header image alt text

`-x, --translateX [translateX]`
: X focal point of the header image

`-y, --translateY [translateY]`
: Y focal point of the header image

`--layout [layout]`
: Layout to use in the header. Allowed values `FullWidthImage,NoImage`. Default `FullWidthImage`

`--textAlignment [textAlignment]`
: How to align text in the header. Allowed values `Center,Left`. Default `Left`

`--showKicker`
: Set, to show the kicker

`--showPublishDate`
: Set, to show the publishing date

`--kicker [kicker]`
: Text to show in the kicker, when showKicker is set

`--authors [authors]`
: Comma-separated list of page authors to show in the header

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified `name` doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

The `showKicker`, `kicker` and `authors` options are based on preview functionality that isn't available on all tenants yet.

## Examples

Reset the page header to default

```sh
m365 spo page header set --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx
```

Reset the page header to default and set authors

```sh
m365 spo page header set --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --authors "steve@contoso.com, bob@contoso.com"
```

Use the specified image focused on the given coordinates in the page header

```sh
m365 spo page header set --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --type Custom --imageUrl /sites/team-a/SiteAssets/hero.jpg --altText 'Sunset over the ocean' --translateX 42.3837520042758 --translateY 56.4285714285714
```

Center the page title in the header and show the publishing date

```sh
m365 spo page header set --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --textAlignment Center --showPublishDate
```
