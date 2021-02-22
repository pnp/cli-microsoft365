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
: Header layout to set on the site. Options: `Standard|Compact|Minimal|Extended`.

`--headerEmphasis [headerEmphasis]`
: Header background to set on the site. Options: `Lightest|Light|Dark|Darkest`.

`--logoAlignment [logoAlignment]`
: Logo position when header layout set to `Extended`. Ignored otherwise. Options: `Left|Center|Right`.

`--footerLayout [footerLayout]`
: Footer layout to set on the site. Options: `Simple|Extended`.

`--footerEmphasis [footerEmphasis]`
: Footer background color to set. Options: `Lightest|Light|Dark|Darkest`.

`--disableMegaMenu [disableMegaMenu]`
: Set to `true` to disable the mega menu and to `false` to enable it. Disabling mega menu results in using the cascading navigation (classic experience). Options: `true|false`.

`--hideTitleInHeader [hideTitleInHeader]`
: Set to `true` to hide the site title in the header and to `false` to show it. Options: `true|false`.

`--disableFooter [disableFooter]`
: Set to `true` to disable the footer on the site and to `false` to enable it. Options: `true|false`.

## Examples

Show site header in compact mode

```sh
m365 spo site chrome set --url https://contoso.sharepoint.com/sites/project-x --headerLayout Compact
```

Show site header in extended mode and display the logo on the right

```sh
m365 spo site chrome set --url https://contoso.sharepoint.com/sites/project-x  --headerLayout Extended --logoAlignment Right
```

Disable the footer on the site

```sh
m365 spo site chrome set --url https://contoso.sharepoint.com/sites/project-x --disableFooter true
```
