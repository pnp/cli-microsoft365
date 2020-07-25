# yammer user get

Retrieves the current user or searches for a user by ID or e-mail

## Usage

```sh
m365 yammer user get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --userId [userId]`
: Retrieve a user by ID

`--email [email]`
: Retrieve a user by e-mail

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

All operations return a single user object. Operations executed with the `email` parameter return an array of user objects.

## Examples
  
Returns the current user

```sh
m365 yammer user get
```

Returns the user with the ID 1496550697

```sh
m365 yammer user get --userId 1496550697
```

Returns an array of users matching the e-mail john.smith@contoso.com

```sh
m365 yammer user get --email john.smith@contoso.com
```

Returns an array of users matching the e-mail john.smith@contoso.com in JSON. The JSON output returns a full user object

```sh
m365 yammer user get --email john.smith@contoso.com --output json
```
