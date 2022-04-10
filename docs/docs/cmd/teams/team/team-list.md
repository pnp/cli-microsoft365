# teams team list

Lists Microsoft Teams teams in the current tenant

## Usage

```sh
m365 teams team list [options]
```

## Options

`-j, --joined`
: Show only joined teams

--8<-- "docs/cmd/_global.md"

## Remarks

You can only see the details or archived status of the Microsoft Teams you are a member of.

## Examples

List all Microsoft Teams in the tenant

```sh
m365 teams team list
```

List all Microsoft Teams in the tenant you are a member of

```sh
m365 teams team list --joined
```
