# Yammer user get

Retrieves the current user or searches for a user by ID or e-mail

## Usage

```sh
yammer user get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--id, --userId [number]`|Retrieve a user by ID
`--email [string]`|Retrieve a user by e-mail
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

All operations return a single user object. Operations executed with the email parameter return an array of user objects.

## Examples
  
Returns the current user

```sh
yammer user get
```

Returns the user with the ID 1496550697

```sh
yammer user get --userId 1496550697
```

Returns an array of users matching the e-mail john.smith@contoso.com

```sh
yammer user get --email john.smith@contoso.com
```

Returns an array of users matching the e-mail john.smith@contoso.com in JSON. The JSON output returns a full user object.
```sh
yammer user get --email john.smith@contoso.com --output json
```