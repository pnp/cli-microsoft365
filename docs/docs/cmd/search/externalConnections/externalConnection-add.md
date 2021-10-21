# search externalConnection add

Add a new external connection to be defined for Microsoft Search

## Usage

```sh
m365 search externalConnection add [options]
```

## Options

`--id [id]`
: Developer-provided unique ID of the connection within the Azure Active Directory tenant. Must be between 3 and 32 characters in length. Must only contain alphanumeric characters. Cannot begin with Microsoft or be one of the following values: None, Directory, Exchange, ExchangeArchive, LinkedIn, Mailbox, OneDriveBusiness, SharePoint, Teams, Yammer, Connectors, TaskFabric, PowerBI, Assistant, TopicEngine, MSFT_All_Connectors. Required.

`--name [name]`
: The display name of the connection to be displayed in the Microsoft 365 admin center. Maximum length of 128 characters. Required.

`--description [description]`
: Description of the connection displayed in the Microsoft 365 admin center. Optional.

`--authorisedAppIds [authorisedAppIds]`
: Comma-separated collection of application IDs for registered Azure Active Directory apps that are allowed to manage the externalConnection and to index content in the externalConnection.

--8<-- "docs/cmd/_global.md"

## Examples

Adds a new external connection with name, displayName and description of test

```sh
m365 search externalconnection add --id test --name test --description test
```

Lists the Microsoft Planner buckets in the Plan _My Plan_ owned by group _My Group_

```sh
m365 search externalconnection add --id test --name test --description test --authorizedAppIds  "00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002"
```
