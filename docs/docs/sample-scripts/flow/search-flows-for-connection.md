# Search flows for connections

Author: [Albert-Jan Schot](https://www.cloudappie.nl/search-flows-connections/)

Search all flows as, an administrator, for a specific search string and report results. This sample allows you to get a report of all flows that are connected to a specific site or list. The `$searchString` can be any value but results are the best when using a GUID or site collection URL.

=== "PowerShell"

    ```powershell
    Write-Output "Retrieving all environments"

    $environments = m365 flow environment list -o json | ConvertFrom-Json
    $searchString = "15f5b014-9508-4941-b564-b4ab1b863a7a" #listGuid
    $path = "exportedflow.json";

    ForEach ($env in $environments) {
        Write-Output "Processing $($env.displayName)..."

        $flows = m365 flow list --environment $env.name --asAdmin -o json | ConvertFrom-Json

        ForEach ($flow in $flows) {
            Write-Output "Processing $($flow.displayName)..."
            m365 flow export --id $flow.name --environment $env.name --format json --path $path

            $flowData = Get-Content -Path $path -ErrorAction SilentlyContinue

            if ($null -ne $flowData) {
                if ($flowData.Contains($searchString)) {
                    Write-Output $($flow.displayName + "contains your search string" + $searchString)
                    Write-Output $flow.id
                }

                Remove-Item $path -Confirm:$false
            }
        }
    }
    ```

Keywords:

- Power Automate
