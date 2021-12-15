# search externalconnection add

Add a new external connection to be defined for Microsoft Search

## Usage

```sh
m365 search externalconnection add [options]
```

## Options

`-i, --id <id>`
: Developer-provided unique ID of the connection within the Azure Active Directory tenant

`-n, --name <name>`
: The display name of the connection to be displayed in the Microsoft 365 admin center. Maximum length of 128 characters

`-d, --description <description>`
: Description of the connection displayed in the Microsoft 365 admin center

`--authorizedAppIds [authorizedAppIds]`
: Comma-separated collection of application IDs for registered Azure Active Directory apps that are allowed to manage the external connection and to index content in the external connection.

--8<-- "docs/cmd/_global.md"

## Remarks

The `id` must be at least 3 and no more than 32 characters long. It can contain only alphanumeric characters, can't begin with _Microsoft_ and can be any of the following values: *None, Directory, Exchange, ExchangeArchive, LinkedIn, Mailbox, OneDriveBusiness, SharePoint, Teams, Yammer, Connectors, TaskFabric, PowerBI, Assistant, TopicEngine, MSFT_All_Connectors*.

## Examples

Adds a new external connection with name and description of test

```sh
m365 search externalconnection add --id MyApp --name "My application" --description "Description of your application"
```

Adds a new external connection with a limited number of authorized apps

```sh
m365 search externalconnection add --id MyApp --name "My application" --description "Description of your application" --authorizedAppIds  "00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002"
```
