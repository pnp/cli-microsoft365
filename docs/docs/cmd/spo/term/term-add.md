# spo term add

Adds taxonomy term

## Usage

```sh
m365 spo term add [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: Name of the term to add

`--termSetId [termSetId]`
: ID of the term set in which to create the term. Specify `termSetId` or `termSetName` but not both

`--termSetName [termSetName]`
: Name of the term set in which to create the term. Specify `termSetId` or `termSetName` but not both

`--termGroupId [termGroupId]`
: ID of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both

`--termGroupName [termGroupName]`
: Name of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both

`-i, --id [id]`
: ID of the term to add

`-d, --description [description]`
: Description of the term to add

`--parentTermId [parentTermId]`
: ID of the term below which the term should be added

`--customProperties [customProperties]`
: JSON string with key-value pairs representing custom properties to set on the term

`--localCustomProperties [localCustomProperties]`
: JSON string with key-value pairs representing local custom properties to set on the term

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

Add taxonomy term with the specified name to the term group and term set specified by their names

```sh
m365 spo term add --name IT --termSetName Department --termGroupName People
```

Add taxonomy term with the specified name to the term group and term set specified by their IDs

```sh
m365 spo term add --name IT --termSetId 8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f --termGroupId 5c928151-c140-4d48-aab9-54da901c7fef
```

Add taxonomy term with the specified name and ID

```sh
m365 spo term add --name IT --id 5c928151-c140-4d48-aab9-54da901c7fef --termSetName Department --termGroupName People
```

Add taxonomy term with custom properties

```sh
m365 spo term add --name IT --termSetName Department --termGroupName People --customProperties '{"Property": "Value"}'
```

Add taxonomy term below the specified term

```sh
m365 spo term add --name IT --parentTermId 5c928151-c140-4d48-aab9-54da901c7fef --termGroupName People
```