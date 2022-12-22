# Authenticate with and call the Microsoft Graph 

Author: [Garry Trinder](https://github.com/garrytrinder)

Calling the Microsoft Graph can be done using the `m365 request` command.

```sh
m365 request --url https://graph.microsoft.com/v1.0/me
```

Another route would be to use shell-specific webrequest tooling. To do this, obtain a new access token for the Microsoft Graph and use it in an HTTP request.

=== "PowerShell"

    ```powershell
    $token = m365 util accesstoken get --resource https://graph.microsoft.com --new --output text
    $me = Invoke-RestMethod -Uri https://graph.microsoft.com/v1.0/me -Headers @{"Authorization"="Bearer $token"}
    $me
    ```

=== "Bash"

    ```bash
    #!/bin/bash

    # requires jq: https://stedolan.github.io/jq/

    token=`m365 util accesstoken get --resource https://graph.microsoft.com --new --output text`
    me=`curl https://graph.microsoft.com/v1.0/me -H "Authorization: Bearer $token"`
    echo $me | jq
    ```
