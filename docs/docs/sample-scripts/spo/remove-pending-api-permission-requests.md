# Remove pending SharePoint API permission requests

Author: [Waldek Mastykarz](https://blog.mastykarz.nl/sample-script-quickly-remove-pending-sharepoint-api-permission-requests/)

When building SharePoint Framework solutions connected to APIs secured with Azure Active Directory,  you'll easily end up with many pending permission requests.

This script helps you to quickly remove pending SharePoint API permission requests.

```javascript tab="JavaScript (Google zx)"
#!/usr/bin/env zx
$.verbose = false;

console.log('Retrieving permission requests...');
const permissionRequests = JSON.parse(await $`m365 spo sp permissionrequest list -o json`);

for (let i = 0; i < permissionRequests.length; i++) {
  const request = permissionRequests[i];
  console.log(`Removing request ${request.Resource}/${request.Scope} (${request.Id})...`);
  try {
    await $`m365 spo sp permissionrequest deny --requestId ${request.Id}`
    console.log(chalk.green('DONE'));
  }
  catch (err) {
    console.error(err.stderr);
  }
}
```

Using [CLI for Microsoft 365](https://aka.ms/cli-m365), the script first retrieves the list of pending SharePoint API permission requests. Then, it iterates through the requests and removes (denies) each one of them using CLI for Microsoft 365. After running this script, your list of pending SharePoint API permission requests will be empty.

This script uses [CLI for Microsoft 365](https://aka.ms/cli-m365) and [Google zx](https://github.com/google/zx).

To run the script, save it to a file with the `.mjs` extension. Next, run the script either by calling `zx remove-permissionrequests.mjs` or `./remove-permissionrequests.mjs` after making the script executable using `chmod +x ./remove-permissionrequests.mjs;`

Keywords:

- SharePoint Online
- API Permissions
