# spo user list

Lists all the users within specific web

## Usage

```sh
m365 spo user list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the web to list the users from

--8<-- "docs/cmd/_global.md"

## Examples

Get list of users in web _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo user list --webUrl https://contoso.sharepoint.com/sites/project-x
```
