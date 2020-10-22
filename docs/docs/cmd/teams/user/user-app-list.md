# teams user app list

List the apps installed in the personal scope of the specified user

## Usage

```sh
m365 teams user app list [options]
```

## Options

`--userId [userId]`
: The ID of the user to get the apps from. Specify `userId` or `userName` but not both.

`--userName [userName]`
: The UPN of the user to get the apps from. Specify `userId` or `userName` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

List the apps installed in the personal scope of the specified user using its ID

```sh
m365 teams user app list --userId 4440558e-8c73-4597-abc7-3644a64c4bce
```

List the apps installed in the personal scope of the specified user using its UPN

```sh
m365 teams user app list --userName admin@contoso.com
```
