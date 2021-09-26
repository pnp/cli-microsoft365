# aad serviceprincipal get

Retrieves a service principal from Azure AD directory

## Usage

```sh
m365 aad serviceprincipal get
```

## Options

`-i, --id [id]`
: The ID of the service principal to retrieve information for. Specify either `id`, `appId` or `name`

`-n, --name [name]`
: The display name of the service principal to retrieve information for. Specify either `id`, `appId` or `name`

`--appId [appId]`
: The appId of the service principal to retrieve information for. Specify either `id` or `appId` or `name`

--8<-- "docs/cmd/_global.md"

## Examples

Get information about the service principal with id _1caf7dcd-7e83-4c3a-94f7-932a1299c843_

```sh
m365 aad serviceprincipal get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c843
```

Get information about service principal with name _ServicePrincipal App_

```sh
m365 aad serviceprincipal get --name "ServicePrincipal App"
```

Get information about service principal with appId _8a2a376d-5f57-4c14-9639-692f841c00bc_

```sh
m365 aad serviceprincipal get --appId "8a2a376d-5f57-4c14-9639-692f841c00bc"
```
