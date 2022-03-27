# search externalconnection remove

Removes the specified new external connection for Microsoft Search

## Usage

```sh
m365 search externalconnection remove [options]
```

## Options

`-i, --id [id]`
: Developer-provided unique ID of the connection within the Azure Active Directory tenant

`-n, --name [name]`
: The display name of the connection to be displayed in the Microsoft 365 admin center. Maximum length of 128 characters

--8<-- "docs/cmd/_global.md"

## Remarks

The `id` must be at least 3 and no more than 32 characters long. It can contain only alphanumeric characters, can't begin with _Microsoft_ and can be any of the following values: *None, Directory, Exchange, ExchangeArchive, LinkedIn, Mailbox, OneDriveBusiness, SharePoint, Teams, Yammer, Connectors, TaskFabric, PowerBI, Assistant, TopicEngine, MSFT_All_Connectors*.

## Examples

Removes external connection with id of TestApp

```sh
m365 search externalconnection remove --id "TestApp"
```

Removes external connection with name of Test App

```sh
m365 search externalconnection remove --name "TestApp"
```
