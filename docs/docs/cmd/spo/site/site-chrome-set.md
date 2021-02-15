# spo site chrome set

Set the chrome header and footer for the specified site

## Usage

```sh
m365 spo site chrome set [options]
```

## Options

`-u, --url <url>`
: URL of the site collection to which you want to change the chrome header/footer

`--headerLayout [headerLayout]`
: Specifies the header layout to set on the site. Options: `Standard|Compact|Minimal|Extended`. Default `Standard`.

`--headerEmphasis [headerEmphasis]`
: Specifies the header its background color to set. Options: `Lightest|Light|Dark|Darkest`. Default `Lightest`.

`--logoAlignment [logoAlignment]`
: When using the `Extended` header, you can set the logo its position. Otherwise this setting will be ignored. Options: `Left|Center|Right`. Default `Left`.

`--footerLayout [footerLayout]`
: Specifies the footer layout to set on the site. Options: `Simple|Extended`. Default `Simple`.

`--footerEmphasis [footerEmphasis]`
: Specifies the footer its background color to set. Options: `Lightest|Light|Dark|Darkest`. Default `Darkest`.

`--disableMegaMenu`
: Specify to disable the mega menu. This results in using the cascading navigation (classic experience).

`--hideTitleInHeader`
: Specify to hide the site title in the header.

`--disableFooter`
: Specify to disable the footer on the site.

## Examples

Update the chrome its header to use the compact style

```sh
m365 spo site chrome set -u https://contoso.sharepoint.com/sites/project-x --headerLayout Compact
```

Update the chrome its header to use the extended style and position the logo to the right

```sh
m365 spo site chrome set -u https://contoso.sharepoint.com/sites/project-x  --headerLayout Extended --logoAlignment Right
```

Disable the footer on the site

```sh
m365 spo site chrome set -u https://contoso.sharepoint.com/sites/project-x --footer
```
