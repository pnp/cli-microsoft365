# aad o365group get

Gets information about the specified Microsoft 365 Group

## Usage

```sh
m365 aad o365group get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The ID of the Microsoft 365 Group to retrieve information for

`--includeSiteUrl`
: Set to retrieve the site URL for the group

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get information about the Microsoft 365 Group with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad o365group get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844
```

Get information about the Microsoft 365 Group with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_ and also retrieve the URL of the corresponding SharePoint site

```sh
m365 aad o365group get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --includeSiteUrl
```