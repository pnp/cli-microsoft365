# spo web add

Create new subsite

## Usage

```sh
m365 spo web add [options]
```

## Options

`-h, --help`
: output usage information

`-t, --title <title>`
: Subsite title

`-d, --description [description]`
: Subsite description

`-u, --webUrl <webUrl>`
: Subsite relative url

`-w, --webTemplate <webTemplate>`
: Subsite template, eg. `STS#0` (Classic team site)

`-p, --parentWebUrl <parentWebUrl>`
: URL of the parent site under which to create the subsite

`-l, --locale [locale]`
: Subsite locale LCID, eg. `1033` for en-US. Default `1033`

`--breakInheritance`
: Set to not inherit permissions from the parent site

`--inheritNavigation`
: Set to inherit the navigation from the parent site

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Create subsite using the _Team site_ template in the _en-US_ locale

```sh
m365 spo web add --title Subsite --description Subsite --webUrl subsite --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --locale 1033
```

Create subsite with unique permissions using the default _en-US_ locale

```sh
m365 spo web add --title Subsite --webUrl subsite --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --breakInheritance
```

Create subsite with the same navigation as the parent site

```sh
m365 spo web add --title Subsite --webUrl subsite --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --inheritNavigation
```