# spo web add

Create subsite

## Usage

```sh
spo web add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-t, --title [title]`|Subsite title
`-d, --description [description]`|Subsite description, optional
`-u, --webUrl [webUrl]`|Subsite relative url
`-w, --webTemplate [webTemplate]`|Subsite template, eg. STS#0 (Classic team site)
`-p, --parentWebUrl [parentWebUrl]`|URL of the parent site under which to create the subsite
`-l, --locale [locale]`|Subsite locale LCID, eg. 1033 for en-US
`--breakInheritance [breakInheritance]`|Set to not inherit permissions from the parent site, optional
`--inheritNavigation [inheritNavigation]`|Set to inherit the navigation from the parent site, optional
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging
 
!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get details of a tenant property, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples


Create subsite

```sh
spo web add --title subsite --description subsite 1 --webUrl "subsite" --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --locale 1033
```

Create subsite with breaking permission inheritance

```sh
spo web add --title subsite --description subsite 1 --webUrl "subsite" --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --locale 1033 --breakInheritance
```

Create subsite with inheriting the navigation

```sh
spo web add --title subsite --description subsite 1 --webUrl "subsite" --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --locale 1033 --inheritNavigation
```

Create subsite with breaking permission inheritance and inheriting the navigation

```sh
spo web add --title subsite --description subsite 1 --webUrl "subsite" --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --locale 1033 --breakInheritance --inheritNavigation
```

## More information

- Creating subsite using REST: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/rest/complete-basic-operations-using-sharepoint-rest-endpoints#creating-a-site-with-rest](https://docs.microsoft.com/en-us/sharepoint/dev/apis/rest/complete-basic-operations-using-sharepoint-rest-endpoints#creating-a-site-with-rest)
