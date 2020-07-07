# aad o365group renew

Renews Office 365 group's expiration

## Usage

```sh
aad o365group renew [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the Office 365 group to renew
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

If the specified _id_ doesn't refer to an existing group, you will get a `The remote server returned an error: (404) Not Found.` error.

## Examples

Renew the Office 365 group with id _28beab62-7540-4db1-a23f-29a6018a3848_

```sh
aad o365group renew --id 28beab62-7540-4db1-a23f-29a6018a3848
```