# search externalconnection schema add

This command allows the administrator to add a schema to a specific external connection for use in Microsoft Search.

## Usage

```sh
m365 search externalconnection schema add [options]
```

## Options

`-i, --externalConnectionId  <externalConnectionId>`
: ID of the External Connection.

`-s, --schema [schema]`
: The schema object to be added.

--8<-- "docs/cmd/_global.md"

## Examples

Adds a new schema to a specific external connection.

```sh
m365 search externalconnection schema add --externalConnectionId 'CliConnectionId' --schema '{"baseType":"microsoft.graph.externalItem","properties":[{"name":"ticketTitle","type":"String","isSearchable":"true","isRetrievable":"true","labels":["title"]},{"name":"priority","type":"String","isQueryable":"true","isRetrievable":"true","isSearchable":"false"},{"name":"assignee","type":"String","isRetrievable":"true"}]}'
```

## Response

The command won't return a response on success.
