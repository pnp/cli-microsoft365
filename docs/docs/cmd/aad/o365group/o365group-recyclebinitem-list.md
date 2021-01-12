# aad o365group recyclebinitem list

Lists Groups from the recycle bin in the current tenant

## Usage

```sh
m365 aad o365group recyclebinitem list [options]
```

## Options

`-d, --displayName [displayName]`
: Lists groups with displayName starting with the specified value

`-m, --mailNickname [mailNickname]`
: Lists groups with mailNickname starting with the specified value

--8<-- "docs/cmd/_global.md"

## Examples

List all deleted Microsoft 365 Groups in the tenant

```sh
m365 aad o365group recyclebinitem list
```

List deleted Microsoft 365 Groups with display name starting with _Project_

```sh
m365 aad o365group recyclebinitem list --displayName Project
```

List deleted Microsoft 365 Groups mail nick name starting with _team_

```sh
m365 aad o365group recyclebinitem list --mailNickname team
```

List deleted Microsoft 365 Groups mail nick name starting with _team_ and with display name starting with _Project_

```sh
m365 aad o365group recyclebinitem list --mailNickname team --displayName Project
```
