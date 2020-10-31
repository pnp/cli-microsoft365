# spo term get

Gets information about the specified taxonomy term

## Usage

```sh
m365 spo term get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id [id]`
: ID of the term to retrieve. Specify `name` or `id` but not both

`-n, --name [name]`
: Name of the term to retrieve. Specify `name` or `id` but not both

`--termGroupId [termGroupId]`
: ID of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both

`--termGroupName [termGroupName]`
: Name of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both

`--termSetId [termSetId]`
: ID of the term set to which the term belongs. Specify `termSetId` or `termSetName` but not both

`--termSetName [termSetName]`
: Name of the term set to which the term belongs. Specify `termSetId` or `termSetName` but not both

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

## Remarks

When retrieving term by its ID, it's sufficient to specify just the ID. When retrieving it by its name however, you need to specify the parent term group and term set using either their names or IDs.

## Examples

Get information about a taxonomy term using its ID

```sh
m365 spo term get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```

Get information about a taxonomy term using its name, retrieving the parent term group and term set using their names

```sh
m365 spo term get --name IT --termGroupName People --termSetName Department
```

Get information about a taxonomy term using its name, retrieving the parent term group and term set using their IDs

```sh
m365 spo term get --name IT --termGroupId 5c928151-c140-4d48-aab9-54da901c7fef --termSetId 8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f
```