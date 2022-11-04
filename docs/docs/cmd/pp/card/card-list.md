# pp card list

Lists Microsoft Power Platform cards in the specified Power Platform environment.

## Usage

```sh
pp card list [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`-a, --asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

List cards in a specific environment.

```sh
m365 pp card list --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
```

List cards in a specific environment as admin.

```sh
m365 pp card list --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --asAdmin
```

## Response

=== "JSON"

    ```json
    [
      {
        "solutionid":"fd140aae-4df4-11dd-bd17-0019b9312238",
        "modifiedon":"2022-10-25T14:44:48Z",
        "_owninguser_value":"4f175d04-b952-ed11-bba2-000d3adf774e",
        "overriddencreatedon":null,
        "ismanaged":false,
        "schemaversion":null,
        "importsequencenumber":null,
        "tags":null,
        "componentidunique":"24205370-bc43-4c5e-b095-16da18f0b1a3",
        "_modifiedonbehalfby_value":null,
        "componentstate":0,
        "statecode":0,
        "name":"Tasks List",
        "versionnumber":4451230,
        "utcconversiontimezonecode":null,
        "cardid":"0eab9392-7354-ed11-bba2-000d3adf774e",
        "publishdate":null,
        "_createdonbehalfby_value":null,
        "_modifiedby_value":"4f175d04-b952-ed11-bba2-000d3adf774e",
        "createdon":"2022-10-25T14:44:48Z",
        "overwritetime":"1900-01-01T00:00:00Z",
        "_owningbusinessunit_value":"b419f090-fe22-ec11-b6e5-000d3ab596a1",
        "hiddentags":null,
        "description":" ",
        "appdefinition":"{\"screens\":{\"main\":{\"template\":{\"type\":\"AdaptiveCard\",\"body\":[{\"type\":\"TextBlock\",\"size\":\"Medium\",\"weight\":\"bolder\",\"text\":\"Your card title goes here\"},{\"type\":\"TextBlock\",\"text\":\"Add and remove element to customize your new card.\",\"wrap\":true}],\"actions\":[],\"$schema\":\"http://adaptivecards.io/schemas/1.4.0/adaptive-card.json\",\"version\":\"1.4\"},\"verbs\":{\"submit\":\"echo\"}}},\"sampleData\":{\"main\":{}},\"connections\":{},\"variables\":{},\"flows\":{}}",
        "statuscode":1,
        "remixsourceid":null,
        "sizes":null,
        "coowners":null,
        "_owningteam_value":null,
        "_createdby_value":"4f175d04-b952-ed11-bba2-000d3adf774e",
        "_ownerid_value":"4f175d04-b952-ed11-bba2-000d3adf774e",
        "publishsourceid":null,
        "timezoneruleversionnumber":null,
        "iscustomizable":{
          "Value":true,
          "CanBeChanged":true,
          "ManagedPropertyLogicalName":"iscustomizableanddeletable"
        },
        "owninguser":{
          "azureactivedirectoryobjectid":"78637d62-e872-4dc9-b7c1-bd161e631682",
          "fullname":"# Nico",
          "systemuserid":"4f175d04-b952-ed11-bba2-000d3adf774e",
          "ownerid":"4f175d04-b952-ed11-bba2-000d3adf774e"
        }
      }
    ]
    ```

=== "Text"

    ```text
    name        cardid                                publishdate          createdon             modifiedon
    ----------  ------------------------------------  -----------          --------------------  --------------------
    Tasks List  0eab9392-7354-ed11-bba2-000d3adf774e  2022-10-30T16:00:00Z 2022-10-25T14:44:48Z  2022-10-25T14:44:48Z
    ```

=== "CSV"

    ```csv
    name,cardid,publishdate,createdon,modifiedon
    Tasks List,0eab9392-7354-ed11-bba2-000d3adf774e,2022-10-30T16:00:00Z,2022-10-25T14:44:48Z,2022-10-25T14:44:48Z
    ```
