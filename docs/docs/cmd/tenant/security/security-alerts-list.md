# tenant security alerts list

Gets the security alerts for a tenant

## Usage

```sh
m365 tenant security alerts list [options]
```

## Options

`--vendor [vendor]`
: The vendor to return alerts for. Possible values `Azure Advanced Threat Protection`, `Azure Security Center`, `Microsoft Cloud App Security`, `Azure Active Directory Identity Protection`, `Azure Sentinel`, `Microsoft Defender ATP`. If omitted, all alerts are returned

--8<-- "docs/cmd/_global.md"

## Examples

Get all security alerts for a tenant

```sh
m365 tenant security alerts list
```

Get security alerts for a vendor with name _Azure Sentinel_

```sh
m365 tenant security alerts list --vendor "Azure Sentinel"
```