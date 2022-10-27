# tenant serviceannouncement health get

Get the health report of a specified service for a tenant

## Usage

```sh
m365 tenant serviceannouncement health get [options]
```

## Options

`-s, --serviceName <serviceName>`
: The service name to retrieve the health report for.

`-i, --issues`
: Return the collection of issues that happened on the service, with detailed information for each issue. Is only returned in JSON output mode.

--8<-- "docs/cmd/_global.md"

## Examples

Get the health report for the service _Exchange Online_

```sh
m365 tenant serviceannouncement health get --serviceName "Exchange Online"
```

Get the health report for the service _Exchange Online_ including the issues of the service

```sh
m365 tenant serviceannouncement health get --serviceName "Exchange Online" --issues
```

## Response

### Standard response

=== "JSON"

    ``` json
    {
      "service": "Exchange Online",
      "status": "serviceDegradation",
      "id": "Exchange"
    }
    ```

=== "Text"

    ``` text
    id     : Exchange
    service: Exchange Online
    status : serviceDegradation
    ```

=== "CSV"

    ``` CSV
    id,status,service
    Exchange,serviceDegradation,Exchange Online
    ```

### `issues` response

When we make use of the option `issues` the response will differ.

=== "JSON"

    ``` json
    {
      "service": "Exchange Online",
      "status": "serviceDegradation",
      "id": "Exchange",
      "issues": [
        {
          "startDateTime": "2022-08-17T18:27:00Z",
          "endDateTime": "2022-08-18T16:06:18Z",
          "lastModifiedDateTime": "2022-08-18T16:46:04.133Z",
          "title": "AdminsÔÇÖ downloaded CSV mailbox usage reports via Microsoft 365 admin center donÔÇÖt contain data prior to August 1, 2022",
          "id": "EX415080",
          "impactDescription": "AdminsÔÇÖ downloaded CSV mailbox usage reports via Microsoft 365 admin center didn't contain data prior to August 1, 2022.",
          "classification": "advisory",
          "origin": "microsoft",
          "status": "serviceRestored",
          "service": "Exchange Online",
          "feature": "Windows Live sign-in issue",
          "featureGroup": "Sign-in",
          "isResolved": true,
          "highImpact": null,
          "details": [],
          "posts": [
            {
              "createdDateTime": "2022-08-17T18:34:54.95Z",
              "postType": "regular",
              "description": {
                "contentType": "html",
                "content": "Title: AdminsÔÇÖ downloaded CSV mailbox usage reports via Microsoft 365 admin center donÔÇÖt contain data prior to August 1, 2022\\\n\nUser impact: AdminsÔÇÖ downloaded CSV mailbox usage reports via Microsoft 365 admin center donÔÇÖt contain data prior to August 1, 2022.\\\n\nCurrent status: We're investigating a potential issue and checking for impact to your organization. We'll provide an update within 30 minutes."
              }
            }
          ]
        }
      ]
    }
    ```

=== "Text"

    ``` text
    id          status              service
    ----------  ------------------  -----------------
    Exchange    serviceDegradation  Exchange Online
    ```

=== "CSV"

    ``` CSV
    id,status,service
    Exchange,serviceDegradation,Exchange Online
    ```
