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
: The resource uri for which the CLI should acquire a token from AAD in order to access
the service.

`-b, --body [body]`
: The request body. Optionally use `@example.json` to load the body from a file.

`-p, --filePath [filePath]`
: The file path to save the response to. This option can be used when downloading files.

--8<-- "docs/cmd/_global.md"

## Remarks

When executing a request, CLI will take care of the very basic configuration, and you'll need to specify all additional information, such as headers, method and body. CLI will take care for you of:

- applying compression and handling throttling,
- setting the `accept` to `application/json` if you don't specify it yourself,
- setting the `authorization` header to the bearer token obtained for the resource determined from the request URL

If you specify the `resource` option, the CLI will try to retrieve a valid token for the resource instead of determining the resource based on the url. The value doesn't have to be a URL. It can be also a URI like `app://<guid>`.

Specify additional headers by typing them as options, for example: `--content-type "application/json"`, `--if-match "*"`, `--x-requestdigest "somedigest"`.

!!! important
    When building the request, depending on the shell you use, you might need to escape all `$` characters in the URL, request headers, and the body. If you don't do it, the shell will treat it as a variable and will remove the following word from the request, breaking the request.

## Examples

Call the SharePoint Rest API using a GET request with a constructed URL containing expands, filters and selects.

```sh
m365 request --url "https://contoso.sharepoint.com/sites/project-x/_api/web/siteusers?\$filter=IsShareByEmailGuestUser eq true&\$expand=Groups&\$select=Title,LoginName,Email,Groups/LoginName" --accept "application/json;odata=nometadata"
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

Call the Microsoft Graph to get a profile photo.

```sh
m365 request --url "https://graph.microsoft.com/beta/me/photo/\$value" --filePath ./profile-pic.jpg
```
