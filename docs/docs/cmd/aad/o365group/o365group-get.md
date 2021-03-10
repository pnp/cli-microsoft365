# aad o365group get

Gets information about the specified Microsoft 365 Group or Microsoft Teams team

## Usage

```sh
m365 aad o365group get [options]
```

## Alias

```sh
m365 teams team get
```

## Options

`-i, --id <id>`
: The ID of the Microsoft 365 Group or Microsoft Teams team to retrieve information for

`--includeSiteUrl`
: Set to retrieve the site URL for the group

--8<-- "docs/cmd/_global.md"

## Examples

Get information about the Microsoft 365 Group with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad o365group get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844
```

Get information about the Microsoft 365 Group with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_ and also retrieve the URL of the corresponding SharePoint site

```sh
m365 aad o365group get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --includeSiteUrl
```

Get information about the Microsoft Teams team with id _2eaf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 teams team get --id 2eaf7dcd-7e83-4c3a-94f7-932a1299c844
```
