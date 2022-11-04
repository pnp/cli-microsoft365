# teams team get

Gets information about the specified Microsoft Teams team

## Usage

```sh
m365 teams team get
```

## Options

`-i, --id [id]`
: The ID of the Microsoft Teams team to retrieve information for. Specify either id or name but not both

`-n, --name [name]`
: The display name of the Microsoft Teams team to retrieve information for. Specify either id or name but not both

--8<-- "docs/cmd/_global.md"

## Examples

Get information about the Microsoft Teams team with id _2eaf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 teams team get --id 2eaf7dcd-7e83-4c3a-94f7-932a1299c844
```

Get information about Microsoft Teams team with name _Team Name_

```sh
m365 teams team get --name "Team Name"
```

## Response

=== "JSON"

    ``` json
    {
      "id": "a40210cd-0060-4b91-aaa1-a44e0853d979",
      "createdDateTime": "2022-10-31T12:50:42.819Z",
      "displayName": "Architecture",
      "description": "Architecture Discussion",
      "internalId": "19:2soiTJiLJmUrSi94Hr23ZwcN9uWFWjE3EGYb5bFsyy41@thread.tacv2",
      "classification": null,
      "specialization": "none",
      "visibility": "public",
      "webUrl": "https://teams.microsoft.com/l/team/19%3a2soiTJiLJmUrSi94Hr23ZwcN9uWFWjE3EGYb5bFsyy41%40thread.tacv2/conversations?groupId=a40210cd-0060-4b91-aaa1-a44e0853d979&tenantId=92e59666-257b-49c3-b1fa-1bae8107f6ba",
      "isArchived": false,
      "isMembershipLimitedToOwners": false,
      "discoverySettings": {
        "showInTeamsSearchAndSuggestions": true
      },
      "summary": null,
      "memberSettings": {
        "allowCreateUpdateChannels": true,
        "allowCreatePrivateChannels": true,
        "allowDeleteChannels": true,
        "allowAddRemoveApps": true,
        "allowCreateUpdateRemoveTabs": true,
        "allowCreateUpdateRemoveConnectors": true
      },
      "guestSettings": {
        "allowCreateUpdateChannels": false,
        "allowDeleteChannels": false
      },
      "messagingSettings": {
        "allowUserEditMessages": true,
        "allowUserDeleteMessages": true,
        "allowOwnerDeleteMessages": true,
        "allowTeamMentions": true,
        "allowChannelMentions": true
      },
      "funSettings": {
        "allowGiphy": true,
        "giphyContentRating": "moderate",
        "allowStickersAndMemes": true,
        "allowCustomMemes": true
      }
    }
    ```

=== "Text"

    ``` text
    classification             : null
    createdDateTime            : 2022-10-31T12:50:42.819Z
    description                : Architecture Discussion
    discoverySettings          : {"showInTeamsSearchAndSuggestions":true}
    displayName                : Architecture
    funSettings                : {"allowGiphy":true,"giphyContentRating":"moderate","allowStickersAndMemes":true,"allowCustomMemes":true}
    guestSettings              : {"allowCreateUpdateChannels":false,"allowDeleteChannels":false}
    id                         : a40210cd-0060-4b91-aaa1-a44e0853d979
    internalId                 : 19:2soiTJiLJmUrSi94Hr23ZwcN9uWFWjE3EGYb5bFsyy41@thread.tacv2
    isArchived                 : false
    isMembershipLimitedToOwners: false
    memberSettings             : {"allowCreateUpdateChannels":true,"allowCreatePrivateChannels":true,"allowDeleteChannels":true,"allowAddRemoveApps":true,"allowCreateUpdateRemoveTabs":true,"allowCreateUpdateRemoveConnectors":true}
    messagingSettings          : {"allowUserEditMessages":true,"allowUserDeleteMessages":true,"allowOwnerDeleteMessages":true,"allowTeamMentions":true,"allowChannelMentions":true}
    specialization             : none
    summary                    : null
    visibility                 : public
    webUrl                     : https://teams.microsoft.com/l/team/19%3a2soiTJiLJmUrSi94Hr23ZwcN9uWFWjE3EGYb5bFsyy41%40thread.tacv2/conversations?groupId=a40210cd-0060-4b91-aaa1-a44e0853d979&tenantId=92e59666-257b-49c3-b1fa-1bae8107f6ba
    ```

=== "CSV"

    ``` text
    id,createdDateTime,displayName,description,internalId,classification,specialization,visibility,webUrl,isArchived,isMembershipLimitedToOwners,discoverySettings,summary,memberSettings,guestSettings,messagingSettings,funSettings
    a40210cd-0060-4b91-aaa1-a44e0853d979,2022-10-31T12:50:42.819Z,Architecture,Architecture Discussion,19:2soiTJiLJmUrSi94Hr23ZwcN9uWFWjE3EGYb5bFsyy41@thread.tacv2,,none,public,https://teams.microsoft.com/l/team/19%3a2soiTJiLJmUrSi94Hr23ZwcN9uWFWjE3EGYb5bFsyy41%40thread.tacv2/conversations?groupId=a40210cd-0060-4b91-aaa1-a44e0853d979&tenantId=92e59666-257b-49c3-b1fa-1bae8107f6ba,,,"{""showInTeamsSearchAndSuggestions"":true}",,"{""allowCreateUpdateChannels"":true,""allowCreatePrivateChannels"":true,""allowDeleteChannels"":true,""allowAddRemoveApps"":true,""allowCreateUpdateRemoveTabs"":true,""allowCreateUpdateRemoveConnectors"":true}","{""allowCreateUpdateChannels"":false,""allowDeleteChannels"":false}","{""allowUserEditMessages"":true,""allowUserDeleteMessages"":true,""allowOwnerDeleteMessages"":true,""allowTeamMentions"":true,""allowChannelMentions"":true}","{""allowGiphy"":true,""giphyContentRating"":""moderate"",""allowStickersAndMemes"":true,""allowCustomMemes"":true}"
    ```
