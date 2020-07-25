# spo term list

Lists taxonomy terms from the given term set

## Usage

```sh
m365 spo term list [options]
```

## Options

`-h, --help`
: output usage information

`--termGroupId [termGroupId]`
: ID of the term group where the term set is located. Specify `termGroupId` or `termGroupName` but not both

`--termGroupName [termGroupName]`
: Name of the term group where the term set is located. Specify `termGroupId` or `termGroupName` but not both

`--termSetId [termSetId]`
: ID of the term set for which to retrieve terms. Specify `termSetId` or `termSetName` but not both

`--termSetName [termSetName]`
: Name of the term set for which to retrieve terms. Specify `termSetId` or `termSetName` but not both

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

List taxonomy terms from the term group and term set with the given name

```sh
m365 spo term list --termGroupName PnPTermSets --termSetName PnP-Organizations
```

List taxonomy terms from the term group and term set with the given ID

```sh
m365 spo term list --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb --termSetId 0e8f395e-ff58-4d45-9ff7-e331ab728bec
```