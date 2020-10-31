# Inventory Flows By Creator  

Author: [Pete Skelly](https://peteskelly.com/lightweight-reports-using-the-office-365-cli-and-jq/)

The [Power Automate Admin Center](https://admin.flow.microsoft.com) provides a list of the Flows in your tenant, but there is no way to easily export Flows from the Flow admin center for governance activities. This script retrieves Flows from the Default Environment and maps creator information from Azure AD to list Flows by owner, state and trigger type.

The `bash` version of this script uses an external file to process owner mapping. This is provided in the jq tab and should be saved to the same folder as the bash script and named `merge.jq`.

!!! attention
    There is a known issue when running scripts that retrieve large amounts of content. See issue [#1266](https://github.com/pnp/cli-microsoft365/issues/1266) for further detail. A best practice is to use a temporary file to enable processing large return sets.

```powershell tab="PowerShell Core"
#!/usr/local/bin/pwsh -File

$DIR = Split-Path $script:MyInvocation.MyCommand.Path
$TMP_DIR = "./tmp"
$TMP_FLOWS = "$TMP_DIR/flows.json"
$FLOWSCSV = "flows.csv"

function CleanDistFolder {
    # Remove the dist folder as needed
    if (Test-Path -Path "$TMP_DIR" -PathType Container) {
        Remove-Item -Path "$TMP_DIR" -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
}

$CURRENT_USER = $(m365 status).Split(':')[1]
Write-Host "Logged in as $CURRENT_USER"

try {
    if (-not (Test-Path -Path "$TMP_DIR" -PathType Container)) {
        Write-Host " Creating $TMP_DIR folder..."
        New-Item -ItemType Directory -Path "$TMP_DIR"
    }

    #Step 1 - Get the default environment
    Write-Host "Querying for default Flow environment..."
    $flowEnvironments = m365 flow environment list --output json | ConvertFrom-Json
    $defaultEnvironment = $flowEnvironments[0].name
    Write-Host "Found default environment $defaultEnvironment, querying Flows..."

    # Step 2 - Get all of the flows using the cli and write flows json to a tmp file 
    # Use a JMESPath query to filter the size of the file. See https://github.com/pnp/cli-microsoft365/issues/1266
    m365 flow list --environment $defaultEnvironment `
        --query '[].{name: name, displayName: properties.displayName,owner: properties.creator.userId, state: properties.state, created: properties.createdTime, lastModified: properties.lastModifiedTime, trigger: properties.definitionSummary.triggers[0].swaggerOperationId,  triggerType: properties.definitionSummary.triggers[0].type }' --asAdmin --output json |
        Out-File "$TMP_FLOWS" -Encoding ASCII
    $flows = Get-Content "$TMP_FLOWS" | ConvertFrom-Json

    #Step 3 - Get a unique list of the flow owners from the tmp file
    Write-Host "Flows found, searching for owner values..."
    $uniqueOwners = $flows.owner | Sort-Object | Get-Unique
    Write-Host "There are $($uniqueOwners.Count) unique Flows owners."
    Write-Host "Building owner information mappings..."

    #Step 4 - map properties.creator.userId's to {name, email} mapping hashtable
    Write-Host "Querying graph for userids..."
    $userMap = @{}
    $uniqueOwners | ForEach-Object {
        Write-Host "Querying graph for userid $_..."
        m365 aad user get --id $_ --output json  | ConvertFrom-Json
    } | ForEach-Object {
        $userMap.Add($_.id, @{
                upn = $_.userPrincipalName
                displayName = $_.displayName
                mail = $_.mail
            }
        )
    }
    # And add the Owner information to each flow entry to get owner name and email  
    Write-Host "Mapping owner information..."
    $flows | ForEach-Object {
         $_ | Add-Member -MemberType NoteProperty -Name "upn" -Value  $userMap[$_.owner].upn
         $_ | Add-Member -MemberType NoteProperty -Name "ownerName" -Value  $userMap[$_.owner].displayName
         $_ | Add-Member -MemberType NoteProperty -Name "ownerMail" -Value  $userMap[$_.owner].mail
     }

    #Step 5 - Create a CSV file with header row, flow information and owner email
    $flows | Export-Csv -Path "$FLOWSCSV" -NoTypeInformation

}
finally {
    CleanDistFolder
}

# if we are on macOS try opening the file with Excel
if ($IsMacOS) {
    $answer = Read-Host -Prompt "Open CSV file in Excel? (y/n)"
    switch ($answer)
     {
       y {
            open -a "/Applications/Microsoft Excel.app" "$DIR/$FLOWSCSV"
        }
       Default {
           Write-Host "Open $DIR/$FLOWSCSV to review report."
        }
     }
}
```

```bash tab="Bash"
#!/usr/bin/env bash
set -e

DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null && pwd )"
TMP_ENVIRONMENTS=./tmp/environments.json
TMP_FLOWS=./tmp/flows.json
TMP_OWNERS=./tmp/owners.json
TMP_MAPPEDFLOWS=./tmp/mappedFlows.json
FLOWSCSV=flows.csv
JQ_MERGE_FILE=merge.jq

function cleanup {
    #clean up the tmp files
    rm -rf tmp
    echo "Cleaned tmp folder..."
}
# Configure trap to call finish whenever EXIT is called to ensure cleanup of tmp
trap cleanup EXIT

CURRENT_USER=$(m365 status --output json | jq '.connectedAs')
echo "Logged in as $CURRENT_USER"

if [[ ! -z tmp ]]; then
    echo "Creating temporary folder for file manipulation..."
    mkdir tmp
fi

#Step 1 - Get the default environment
echo "Querying for default Flow environment..."
DEFAULT_ENVIRONMENT=$(m365 flow environment list --output json | jq -r '.[] | select(.name | contains("'"Default"'")) | .name')
echo "Found default environment $DEFAULT_ENVIRONMENT, querying Flows..."

#Step 2 - Get all of the flows using the cli and write flows json to a tmp file
#See https://github.com/pnp/cli-microsoft365/issues/1266 for temp file usage reason
m365 flow list --environment $DEFAULT_ENVIRONMENT --asAdmin --output json > $TMP_FLOWS

#Step 3 - Get a unique list of the flow owners from the tmp file
echo "Flows found, searching for owner values..."
uniqueOwners=$(cat $TMP_FLOWS | jq -r 'map({userId: .properties.creator.userId}) | unique | .[] | .userId') 

#Get the owner count and loop to call Microsoft Graph and build owner mapping file  
ownerCount=$(cat $TMP_FLOWS | jq -r 'map({userId: .properties.creator.userId}) | unique | length') 

echo "There are $ownerCount unique Flows owners."
echo "Building owner information json mapping file..."
echo "[" > $TMP_OWNERS  
i=0
for ownerId in $uniqueOwners; do
    echo "Querying graph for userid $ownerId..."
    echo $(m365 aad user get --id $ownerId --output json) >> $TMP_OWNERS
    if [[ $i -lt $ownerCount-1 ]]; then
        echo "," >> $TMP_OWNERS
    fi
    i=$(expr $i + 1)  
done
echo "]" >> $TMP_OWNERS  

#Step 4 - Use a jq module file to create a map of the creator.usedId's to {name, email}
echo "Mapping owners information..."
jq -n --argfile flows $TMP_FLOWS --argfile owners $TMP_OWNERS -f $JQ_MERGE_FILE >> $TMP_MAPPEDFLOWS

#Step 5 - Create a CSV file with header row, flow information and owner email
echo "Building CSV file..."
jq -r '["FlowID", "DisplayName", "State", "Created", "LastModified", "Owner", "OwnerName", "OwnerMail", "Upn", "Trigger", "TriggerType"], (.[] | [.name, .properties.displayName, .properties.state, .properties.createdTime, .properties.lastModifiedTime, .properties.creator.userId, .properties.creator.displayName, .properties.creator.mail, .properties.creator.userPrincipalName, .properties.definitionSummary.triggers[0].swaggerOperationId, .properties.definitionSummary.triggers[0].type]) | @csv' $TMP_MAPPEDFLOWS > $FLOWSCSV

# if we are on macOS try opening the file with Excel
if [[ "$OSTYPE" == "darwin"* ]]; then
    echo "Open CSV file in Excel? (y/n)?"
    read answer
    if [ "$answer" != "${answer#[Yy]}" ] ;then
        open -a /Applications/Microsoft\ Excel.app $DIR/$FLOWSCSV
    else
        echo "Open $DIR/file.csv to review report."
    fi
fi
```

```bash tab="jq"
# Create a dictionary based on the $owner.id property from the owners array parameter
($owners | map(select(.id != null)) | map( {(.id): {displayName, userPrincipalName, mail}}) | add) as $dict
# Output each flow, append owner information from each entry using flow creator.userId property as the key
| $flows |.[].properties.creator |= . + $dict[.userId]
```

Keywords:

- Power Automate
- Azure Active Directory