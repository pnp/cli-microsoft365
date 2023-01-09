# pp card get

Gets a specific Microsoft Power Platform card in the specified Power Platform environment

## Usage

```sh
pp card get [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`-i, --id [id]`
: The id of the card. Specify either `id` or `name` but not both.

`-n, --name [name]`
: The name of the card. Specify either `id` or `name` but not both.

`--asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

Get a specific card in a specific environment based on name

```sh
m365 pp card get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --name "CLI 365 Card"
```

Get a specific card in a specific environment based on name as admin

```sh
m365 pp card get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --name "CLI 365 Card" --asAdmin
```

Get a specific card in a specific environment based on id

```sh
m365 pp card get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --id "408e3f42-4c9e-4c93-8aaf-3cbdea9179aa"
```

Get a specific card in a specific environment based on id as admin

```sh
m365 pp card get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --id "408e3f42-4c9e-4c93-8aaf-3cbdea9179aa" --asAdmin
```

## Response

=== "JSON"

    ```json
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
    ```

=== "Text"

    ```text
    cardid     : 0eab9392-7354-ed11-bba2-000d3adf774e
    createdon  : 2022-10-25T14:44:48Z
    modifiedon : 2022-10-25T14:44:48Z
    name       : Tasks List
    publishdate: 2022-10-30T16:00:00Z
    ```

=== "CSV"

    ```csv
    name,cardid,publishdate,createdon,modifiedon
    Tasks List,0eab9392-7354-ed11-bba2-000d3adf774e,2022-10-30T16:00:00Z,2022-10-25T14:44:48Z,2022-10-25T14:44:48Z
    ```

=== "Markdown"

    ```md
    # pp card get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --id "0eab9392-7354-ed11-bba2-000d3adf774e"

    Date: 9/1/2023

    ## Tasks List (undefined)

    Property | Value
    ---------|-------
    solutionid | fd140aae-4df4-11dd-bd17-0019b9312238
    modifiedon | 2022-10-25T14:44:48Z
    \_owninguser\_value | 4f175d04-b952-ed11-bba2-000d3adf774e
    overriddencreatedon | null
    ismanaged | false
    schemaversion | null
    tags | null
    importsequencenumber | null
    componentidunique | 24205370-bc43-4c5e-b095-16da18f0b1a3
    \_modifiedonbehalfby\_value | null
    componentstate | 0
    statecode | 0
    name | Tasks List
    versionnumber | 4451230
    utcconversiontimezonecode | null
    cardid | 0eab9392-7354-ed11-bba2-000d3adf774e
    publishdate | null
    \_createdonbehalfby\_value | null
    \_modifiedby\_value | 4f175d04-b952-ed11-bba2-000d3adf774e
    createdon | 2022-10-25T14:44:48Z
    overwritetime | 1900-01-01T00:00:00Z
    \_owningbusinessunit\_value | b419f090-fe22-ec11-b6e5-000d3ab596a1
    hiddentags | null
    description |
    appdefinition | {"screens":{"main":{"template":{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"bolder","text":"Your card title goes here"},{"type":"TextBlock","text":"Add and remove element to customize your new card.","wrap":true}],"actions":[],"$schema":"http://adaptivecards.io/schemas/1.4.0/adaptive-card.json","version":"1.4"},"verbs":{"submit":"echo"}}},"sampleData":{"main":{}},"connections":{},"variables":{},"flows":{}}
    statuscode | 1
    remixsourceid | null
    sizes | null
    \_createdby\_value | 4f175d04-b952-ed11-bba2-000d3adf774e
    \_owningteam\_value | null
    coowners | null
    \_ownerid\_value | 4f175d04-b952-ed11-bba2-000d3adf774e
    publishsourceid | null
    timezoneruleversionnumber | null
    iscustomizable | {"Value":true,"CanBeChanged":true,"ManagedPropertyLogicalName":"iscustomizableanddeletable"}
    owninguser | {"azureactivedirectoryobjectid":"78ccf530-bbf0-47e4-aae6-da5f8c6fb142","fullname":"# Nico","systemuserid":"4f175d04-b952-ed11-bba2-000d3adf774e","ownerid":"4f175d04-b952-ed11-bba2-000d3adf774e"}
    ```
