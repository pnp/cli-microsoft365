# Empty the tenant recycle bin

Author: [Laura Kokkarinen](https://laurakokkarinen.com/does-it-spark-joy-powershell-scripts-for-keeping-your-development-environment-tidy-and-spotless/#empty-the-tenant-recycle-bin)

Your deleted modern SharePoint sites are not going to disappear from the UI before they have been removed from the tenant recycle bin. You can either wait for three months, delete them manually via the SharePoint admin center, or run the CLI for Microsoft 365 script below.

```powershell tab="PowerShell"
$deletedSites = m365 spo tenant recyclebinitem list -o json | ConvertFrom-Json
$deletedSites | Format-Table Url

if ($deletedSites.Count -eq 0) { break }

Read-Host -Prompt "Press Enter to start deleting (CTRL + C to exit)"

$progress = 0
$total = $deletedSites.Count

foreach ($deletedSite in $deletedSites)
{
  $progress++
  Write-Host $progress / $total":" $deletedSite.Url
  m365 spo tenant recyclebinitem remove -u $deletedSite.Url --confirm
}
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/

deletedsites=( $(m365 spo tenant recyclebinitem list -o json | jq -r '.[].Url') )

if [ ${#deletedsites[@]} = 0 ]; then
  exit 1
fi

printf '%s\n' "${deletedsites[@]}"
echo "Press Enter to start deleting (CTRL + C to exit)"
read foo

progress=0
total=${#deletedsites[@]}

for deletedsite in "${deletedsites[@]}"; do
  ((progress++))
  printf '%s / %s:%s\n' "$progress" "$total" "$deletedsite"
  m365 spo tenant recyclebinitem remove -u $deletedsite --confirm
done
```

Keywords:

- SharePoint Online
