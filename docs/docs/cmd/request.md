# request

Executes the specified web request using CLI for Microsoft 365

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
the service.

`-b, --body [body]`
: The request body. Optionally use `@example.json` to load the body from a file. 

`-a, --accept [accept]`
: A convenience option to set the `accept` header of the request. The default to `application/json`. 

--8<-- "docs/cmd/_global.md"

## Remarks

The request will be issued as bare as possible, meaning you are responsible for all request information, such as headers, method and body. The only exception is the `authorization` header. By default, the command will try to retrieve a valid token for the API you are executing a request against based on the `url` option. 

If you specify the `resource` option, the CLI will try to retrieve a valid token for the resource instead of determining the resource based on the url. 

Specify additional headers by typing them as options, for example: `--content-type "application/json"`, `--if-match "*"`, `--x-requestdigest "somedigest"`

## Examples

Call the SharePoint Rest API using a GET request with a constructed URL containing expands, filters and selects.

```sh
m365 request --url "https://contoso.sharepoint.com/sites/project-x/_api/web/siteusers?$filter=IsShareByEmailGuestUser eq true&$expand=Groups&$select=Title,LoginName,Email,Groups/LoginName" --accept "application/json;odata=nometadata"
```

Call the Microsoft Graph beta endpoint using a GET request.

```sh
m365 request --url "https://graph.microsoft.com/beta/me"
```

Call the SharePoint API to retrieve a form digest.

```sh
m365 request --method post --url "https://contoso.sharepoint.com/sites/project-x/_api/contextinfo"
```

Call the SharePoint API to update a site title.

```sh
m365 request --method post --url "https://contoso.sharepoint.com/sites/project-x/_api/web" --body '{ "Title": "New title" }' --content-type "application/json" --x-http-method "PATCH"
```
