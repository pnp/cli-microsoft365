# tenant serviceannouncement message list

Gets all service update messages for the tenant.

## Usage

```sh
m365 tenant serviceannouncement message list [options]
```

## Options

`-s, --service [service]`
: Retrieve service update messages for the particular service. If not provided, retrieves messages for all services

--8<-- "docs/cmd/_global.md"

## Examples

Get service update messages of all services in Microsoft 365

```sh
m365 tenant serviceannouncement message list
```

Get service update messages for Microsoft Teams

```sh
m365 tenant serviceannouncement message list --service "Microsoft Teams"
```

## Response

=== "JSON"

    ```json
    [
      {
        "startDateTime": "2021-07-08T00:37:37Z",
        "endDateTime": "2022-12-09T07:00:00Z",
        "lastModifiedDateTime": "2022-06-07T20:21:06.713Z",
        "title": "(Updated) Microsoft Lists: Custom list templates",
        "id": "MC267581",
        "category": "planForChange",
        "severity": "normal",
        "tags": [
          "Updated message",
          "New feature",
          "User impact",
          "Admin impact"
        ],
        "isMajorChange": false,
        "actionRequiredByDateTime": null,
        "services": [
          "SharePoint Online"
        ],
        "expiryDateTime": null,
        "hasAttachments": false,
        "viewPoint": null,
        "details": [],
        "body": {
          "contentType": "html",
          "content": "<p>Updated June 7, 2022: We have updated the rollout timeline below. Thank you for your patience.</p><p>This new feature will support the addition of custom list templates from your organization alongside the ready-made templates Microsoft provides to make it easy to get started tracking and managing information.</p> \\\n<p>[Key points]</p> \\\n<ul> \\\n<li>Microsoft 365 <a href=\"https://www.microsoft.com/microsoft-365/roadmap?filters=&amp;searchterms=70753\" target=\"_blank\">Roadmap ID: 70753</a></li> \\\n<li>Timing:<ul><li>Targeted release (entire org): Complete</li><li>Standard release: will roll out in mid-September (previously mid-May) and be complete by early November (previously mid-June)</li></ul></li> \\\n<li>Roll-out: tenant level </li> \\\n<li>Control type: user control / admin control</li> \\\n<li>Action: review, assess and educate</li></ul><p>[How this will affect your organization]</p><p>This feature will give organizations the ability to create their own custom list templates with custom formatting and schema. It will also empower organizations to create repeatable solutions within the same Microsoft Lists infrastructure (including list creation in SharePoint, Teams, and within the Lists app itself).</p><p>End-user impact:</p>\\\n<p>Visual updates to the list creation dialog and the addition of a<i> From your organization</i> tab when creating a new list. This new tab is where your custom list templates appear alongside the ready-made templates from Microsoft.</p>\\\n<p><img src=\"https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE4P81n?ver=c93f\" alt=\"Your custom list templates along with Microsoft ready-made templates\" width=\"550\"><br>\\\nAdmin impact:</p>\\\n<p>Custom list templates can only be uploaded by a SharePoint administrator for Microsoft 365 by using PowerShell cmdlets. For consistency, the process of defining and uploading custom list templates is like the custom site templates experience.</p><p>To define and upload custom list templates, admins will use the following site template PowerShell cmdlets:</p><ul><li>Use the <a href=\"https://docs.microsoft.com/powershell/module/sharepoint-online/get-spositescriptfromlist?view=sharepoint-ps\" target=\"_blank\">Get-SPOSiteScriptFromList</a> cmdlet to extract the site script from any list</li><li>Run <a href=\"https://docs.microsoft.com/powershell/module/sharepoint-online/add-spositescript?view=sharepoint-ps\" target=\"_blank\">Add-SPOSiteScript</a> and <b style=\"\">Add-SPOListDesign</b> to add the custom list template to your organization.</li><li>Scope who sees the template by using <a href=\"https://docs.microsoft.com/powershell/module/sharepoint-online/grant-spositedesignrights?view=sharepoint-ps\" target=\"_blank\">Grant-SPOSiteDesignRights</a>  (Optional).</li></ul><p>The visual updates for this feature will be seen by end-users in the updated user interface (UI) when creating a list.</p><p>The <i>From your organization</i> tab will be empty until your organization defines and publishes custom list templates.</p>\\\n<p><img src=\"https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE4P81t?ver=70be\" alt=\"Your custom list templates along with Microsoft ready-made templates\" width=\"550\"></p><p>[What you need to do to prepare]</p><p>You might want to notify your users about this new capability and update your training and documentation as appropriate.</p>\\\n<p>Learn more:</p><ul><li><a href=\"https://docs.microsoft.com/sharepoint/dev/declarative-customization/site-design-overview\" target=\"_blank\">PowerShell Cmdlets documentation for custom list templates</a></li><li> <a href=\"https://docs.microsoft.com/sharepoint/lists-custom-template\" target=\"_blank\">Creating custom list templates</a></li></ul>"
        }
      }
    ]
    ```

=== "Text"

    ```text
    id        title
    --------  ------------------------------------------------
    MC267581  (Updated) Microsoft Lists: Custom list templates
    ```

=== "CSV"

    ```csv
    id,title
    MC267581,(Updated) Microsoft Lists: Custom list templates
    ```

=== "Markdown"

    ```md
    # tenant serviceannouncement message list

    Date: 5/24/2022

    ## (Updated) Microsoft Lists: Custom list templates (MC267581)

    Property | Value
    ---------|-------
    startDateTime | 2021-07-08T00:37:37Z
    endDateTime | 2022-12-09T07:00:00Z
    lastModifiedDateTime | 2023-04-11T19:33:45.553Z
    title | (Updated) Microsoft Lists: Custom list templates
    id | MC267581
    category | planForChange
    severity | normal
    isMajorChange | false
    hasAttachments | false
    ```
 
## More information

- List serviceAnnouncement messages: [https://docs.microsoft.com/en-us/graph/api/serviceannouncement-list-messages](https://docs.microsoft.com/en-us/graph/api/serviceannouncement-list-messages)
