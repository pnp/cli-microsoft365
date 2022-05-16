# spo userprofile get

Get SharePoint user profile properties for the specified user

## Usage

```sh
spo userprofile get [options]
```

## Options

`-u, --userName <userName>`
: Account name of the user

--8<-- "docs/cmd/_global.md"

## Remarks

You have to have tenant admin permissions in order to use this command to get profile properties of other users.

## Examples

 Get SharePoint user profile for the specified user

```sh
m365 spo userprofile get --userName 'john.doe@mytenant.onmicrosoft.com'
```
