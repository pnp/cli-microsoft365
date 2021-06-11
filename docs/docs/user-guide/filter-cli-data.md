# Filter CLI data using JMESPath queries

CLI for Microsoft 365 supports filtering, sorting and querying data using [JMESPath](http://jmespath.org/) queries. By specifying the `--query` option on each command you can create complex queries.

There are two types of data returned by the CLI for Microsoft 365 when retrieving data as JSON. In most cases it returns an array of items, but some of the older commands the response is encapsulated in a `value` object. For both scenario's you can use JMESPath to filter, but the queries are a bit different.

## Testing JMESPath queries

You can test your queries using the [JMESPath](http://jmespath.org/) interactive homepage. You can execute a CLI for Microsoft 365 command, get the JSON response and paste it on the homepage and test your queries from there if you do not want to test them while writing scripts. It is a great way to learn what's possible!

## Basic array filters

Let's start with a basic command and return some results using the following command: `m365 spo site classic list --output json`. To simplify the testing most of the properties are removed, but the result would look similarly to:

```json
[{
    "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
    "AllowDownloadingNonWebViewableFiles": true,
    "AllowEditing": false,
    "Title": "Demo 1",
    "Status": "Active",
    "StorageMaximumLevel": 26214400,
    "StorageQuotaType": null
 },
 {
    "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
    "AllowDownloadingNonWebViewableFiles": false,
    "AllowEditing": false,
    "Title": "A Demo 2",
    "Status": "Active",
    "StorageMaximumLevel": 26214400,
    "StorageQuotaType": null
 },
 {
    "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
    "AllowDownloadingNonWebViewableFiles": true,
    "AllowEditing": false,
    "Title": "Sample 1",
    "Status": "Active",
    "StorageMaximumLevel": 26214400,
    "StorageQuotaType": null
}]
```

Using JMESPath queries you can do basic filtering as well as more complex scenario's like `starts_with` or `ends_with`.

- `[?Title == 'Demo 1']` returns only the **first item** from the array as it matches on the title **Demo 1**
- `[?contains(Title, 'Demo')]` would return the **first two items** as it matches the title on the word **Demo**
- `[?contains(*, 'Demo 1')]` would return any item in the array where the value of any property is **Demo 1**, currently only the **first item**.
- `[?starts_with(Title, 'Demo')]` would only return the **first item** as it filters the title to start with **Demo**
- `[?ends_with(Title, '1')]` returns the **first and last item**, as the title ends with a **1**
- `[?contains(Title, 'Demo') && AllowDownloadingNonWebViewableFiles]` returns only the **first item** as it **combines a title filter, and a check on AllowDownloadingNonWebViewableFiles** two filters.

Besides filtering you can also scope what will be returned as a result:

- `[*].Title` returns only the Title for all items.
- `[*].{Title: Title}` returns all items as array with a Title property.

Or you can combine both a filter query and scope the results:

- `[?contains(Title, 'Demo') && AllowDownloadingNonWebViewableFiles].Title`  returns only the **first item title**

## Other array filters

Some commands in the CLI still return their data wrapped in a `value` object. You can still use JMESPath for those, but a query will look slightly different. Executing the following command `m365 spo user list --webUrl https://contoso.sharepoint.com/ --output json` will return a dataset that is similar to the sample below:

```json
{
  "value": [
    {
      "Id": 7,
      "LoginName": "i:0#.f|membership|garth@contoso.nl",
      "Title": "Garth North",
      "PrincipalType": 1,
      "Email": "garth@contoso.nl",
      "IsEmailAuthenticationGuestUser": false,
      "IsShareByEmailGuestUser": false,
      "IsSiteAdmin": true,
      "UserId": {
        "NameId": "xxxx",
        "NameIdIssuer": "urn:federation:microsoftonline"
      },
      "UserPrincipalName": "garth@contoso.nl"
    },
    {
      "Id": 2,
      "LoginName": "i:0#.f|membership|admin@contoso.nl",
      "Title": "Admin",
      "Email": "Admin@contoso.nl",
      "Expiration": "",
      "IsEmailAuthenticationGuestUser": false,
      "IsShareByEmailGuestUser": false,
      "IsSiteAdmin": true,
      "UserId": {
        "NameId": "xxxx",
        "NameIdIssuer": "urn:federation:microsoftonline"
      },
      "UserPrincipalName": "admin@contoso.nl"
    }
  ]
}
```

- `value[?Title == 'Garth North']` returns only the **first item** from the array as it matches on **Garth North**
- `value[?contains(Email, 'Contoso')]` would return the **all items** as it matches the **Email** on the word **Contoso**
- `value[?contains(*, 'North')]` would return any item in the array where the value of **any property is North**, currently only the **first item**.
- `[?starts_with(Title, 'Garth')]` would only return the **first item** as it filters the **title to start with Garth**
- `[?ends_with(UserPrincipalName, '.nl')]` returns the **all items**, as the **title ends with a .nl**
- `[?contains(Title, 'Garth') && IsSiteAdmin]` returns only the **first item** as it **combines a title filter, and a check on IsSiteAdmin** two filters.

!!! important
    All JMESPath queries are case sensitive

For complete list of filter options check out the [JMESPath Examples](https://jmespath.org/examples.html).
