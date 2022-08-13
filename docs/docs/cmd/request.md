# request

Invoke a custom request at a Microsoft 365 API

## Usage

```sh
m365 request [options]
```

## Options

`-u, --url <url>`
: The request URL. 

`-m, --method [method]`
: The HTTP request method. Accepted values are `get, post, put, patch, delete, head, options`. The default value is `get`.

`-r, --resource [resource]`
: The resource url for which the CLI should acquire a token from AAD in order to access 
the service. The token will be placed in the Authorization header. By default, the CLI can figure this out based on the `--url` argument.

`-b, --body [body]`
: The request body. Optionally use `@example.json` to load the body from a file. 

`-h, --headers [headers]`
: A JSON string containing optional header values. Optionally use `@example.json` to load the JSON from a 
file.

`-a, --accept [accept]`
: A convenience option to set the Accept header of the request.

`-c, --contentType [contentType]`
: A convenience option to set the Content-Type header of the request.

--8<-- "docs/cmd/_global.md"

## Remarks

The request will be issued as bare as possible, meaning you are responsible for all request information, such as headers, method and body. The only exception is the `Authorization` header. The command will try to retrieve a valid token for the API you are executing a request against based on the `url` or `resource` options.   

If the `Accept` value is not set on the option or in the headers, it will default to `application/json`. 

## Examples

Call the SharePoint Rest API using a GET request with a constructed URL containing expands, filters and selects.

```sh
m365 request --url "http://contoso.sharepoint.com/sites/project-x/_api/web/siteusers?$filter=IsShareByEmailGuestUser eq true&$expand=Groups&$select=Title,LoginName,Email,Groups/LoginName" --accept "application/json"
```

Call the Microsoft Graph beta endpoint using a GET request.

```sh
m365 request --url "https://graph.microsoft.com/beta/me"
```

Call the SharePoint API to retrieve a form digest.

```sh
m365 request --method post --url "http://contoso.sharepoint.com/sites/project-x/_api/contextinfo" --accept "application/json"
```

Call the SharePoint API to update a site title.

```sh
m365 request --method post --url "http://contoso.sharepoint.com/sites/project-x/_api/web" --headers '{"X-HTTP-Method": "PATCH"}' --body '{ "Title": "New title" }' --contentType "application/json"
```
