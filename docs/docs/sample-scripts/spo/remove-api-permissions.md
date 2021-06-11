# Remove SharePoint API permissions

Author: [Waldek Mastykarz](https://blog.mastykarz.nl/sample-script-quickly-remove-sharepoint-api-permissions/)

When building SharePoint Framework solutions connected to APIs secured with Azure Active Directory, you might need to clear the list of granted API permissions.

This script helps you to quickly remove SharePoint API permissions.

=== "JavaScript (Google zx)"

    ```javascript
    #!/usr/bin/env zx
    $.verbose = false;

    console.log('Retrieving granted API permissions...');
    const apiPermissions = JSON.parse(await $`m365 spo sp grant list -o json`);

    for (let i = 0; i < apiPermissions.length; i++) {
      const permission = apiPermissions[i];
      console.log(`Removing permission ${permission.Resource}/${permission.Scope} (${permission.ObjectId})...`);
      try {
        await $`m365 spo serviceprincipal grant revoke --grantId ${permission.ObjectId}`
        console.log(chalk.green('DONE'));
      }
      catch (err) {
        console.error(err.stderr);
      }
    }
    ```

Using [CLI for Microsoft 365](https://aka.ms/cli-m365), the script first retrieves the list of granted API permissions. Then, it iterates through them and removes (revokes) each one of them using CLI for Microsoft 365. After running this script, your list of SharePoint API permissions will be empty.

This script uses [CLI for Microsoft 365](https://aka.ms/cli-m365) and [Google zx](https://github.com/google/zx).

To run the script, save it to a file with the `.mjs` extension. Next, run the script either by calling `zx remove-apipermissions.mjs` or `./remove-apipermissions.mjs` after making the script executable using `chmod +x ./remove-apipermissions.mjs;`

Keywords:

- SharePoint Online
- API Permissions
