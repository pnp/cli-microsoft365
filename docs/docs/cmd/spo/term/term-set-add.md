# spo term set add

Adds taxonomy term set

## Usage

```sh
m365 spo term set add [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: Name of the term set to add

`--termGroupId [termGroupId]`
: ID of the term group in which to create the term set. Specify `termGroupId` or `termGroupName` but not both

`--termGroupName [termGroupName]`
: Name of the term group in which to create the term set. Specify `termGroupId` or `termGroupName` but not both

`-i, --id [id]`
: ID of the term set to add

`-d, --description [description]`
: Description of the term set to add

`--customProperties [customProperties]`
: JSON string with key-value pairs representing custom properties to set on the term set

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Add taxonomy term set to the term group specified by ID

```sh
m365 spo term set add --name PnP-Organizations --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```

Add taxonomy term set to the term group specified by name. Create the term set with the specified ID

```sh
m365 spo term set add --name PnP-Organizations --termGroupName PnPTermSets --id aa70ede6-83d1-466d-8d95-30d29e9bbd7c
```

Add taxonomy term set and set its description

```sh
m365 spo term set add --name PnP-Organizations --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb --description 'Contains a list of organizations'
```

Add taxonomy term set and set its custom properties

```sh
m365 spo term set add --name PnP-Organizations --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb --customProperties '`{"Property":"Value"}`'
```