# aad policy list

Returns policies from Azure AD

## Usage

```sh
m365 aad policy list [options]
```

## Options

`-p, --policyType [policyType]`
: The type of policies to return. Allowed values `activityBasedTimeout`,`authorization`,`claimsMapping`,`homeRealmDiscovery`,`identitySecurityDefaultsEnforcement`,`tokenIssuance`,`tokenLifetime`. If omitted, all policies are returned

--8<-- "docs/cmd/_global.md"

## Examples

Returns all policies from Azure AD

```sh
m365 aad policy list
```

Returns claim-mapping policies from Azure AD

```sh
m365 aad policy list --policyType "claimsMapping"
```

## More information

- Microsoft Graph Azure AD policy overview: [https://docs.microsoft.com/en-us/graph/api/resources/policy-overview?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/policy-overview?view=graph-rest-1.0)
