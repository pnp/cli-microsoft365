# Planner migration to SharePoint list

Author: [Albert-Jan Schot](https://www.cloudappie.nl/migrate-planner-plan-list-climicrosoft365/)

Use the CLI for Microsoft 365 to migrate an existing plan to a SharePoint Online List with this sample. You can specify the planner plan that you want to migrate and the script will generate a new list for you with the required fields and views. The current sample does migrate all tasks however it skips the following information:

- Categories
- Comments made on tasks
- Attachments

=== "PowerShell"

    ```powershell
    param (
        [Parameter(Mandatory = $true, HelpMessage = "URL of the target site", Position = 0)]
        [string]$SiteUrl,
        [Parameter(Mandatory = $true, HelpMessage = "Groupname or Planner Plan name", Position = 1)]
        [string]$PlanGroupName,
        [Parameter(HelpMessage = "Show progress messages", Position = 2)]
        [switch]$ShowProgress,
        [Parameter(HelpMessage = "Skip List creation (for running multiple migrations without creating the list)", Position = 3)]
        [switch]$SkipListCreation
    )

    $m365Status = m365 status

    if ($m365Status -match "Logged Out") {
        # Connection to Microsoft 365
        m365 login
    }

    $plans = m365 planner plan list --ownerGroupName $PlanGroupName | ConvertFrom-Json
    Write-Host "Found $($plans.length) plans to migrate"

    foreach ($plan in $plans) {
        $migrationTasks = @()
        $plannerBuckets = m365 planner bucket list --planId $plan.id | ConvertFrom-Json
        $plannerTasks = m365 planner task list --planId $plan.id | ConvertFrom-Json

        foreach ($task in $plannerTasks) {
            $taskDetails = m365 planner task get --id $task.id --query '{description: description, checklist: checklist.*.{isChecked: isChecked, title: title} }' | ConvertFrom-Json

            $assignedUsers = @()
            foreach ($userId in $($task.assignments | ForEach-Object { $($_).PSObject.Properties.Name })) {
                $assignedUsers += "{'Key':'i:0#.f|membership|$(m365 aad user get --id $userId --query "userPrincipalName" | ConvertFrom-Json)'}"
            }

            $checklist = $null;

            foreach ($item in $taskDetails.checklist) {
                if ($item.isChecked) {
                    $checklist += "[x] " + $item.title + "<br>";
                }
                else {
                    $checklist += "[ ] " + $item.title + "<br>";
                }
            }

            $migrationTasks += [pscustomobject][ordered]@{
                Title       = $task.title
                Bucket      = ($plannerBuckets | Where-Object { $_.id -eq $task.bucketId }).Name
                Progress    = $task.percentComplete
                Priority    = $task.priority
                Description = ($task.hasDescription ? $taskDetails.description : ' ')
                StartDate   = $task.startDateTime
                DueDate     = $task.dueDateTime
                Checklist   = $checklist
                AssignedTo  = $($assignedUsers ? "[$($assignedUsers -join ",")]" : $null)
            }
        }

        Write-Host "`nFound $($plannerBuckets.length) buckets and $($plannerTasks.length) tasks to migrate for the plan $($plan.title)"

        if ($plannerTasks.length -gt 0) {
            if ($false -eq $SkipListCreation) {
                if ($ShowProgress) {
                    Write-Host "Setting up List"
                }

                $list = m365 spo list add --title $plan.title --baseTemplate GenericList --webUrl $SiteUrl | ConvertFrom-Json

                $bucketOptions = $plannerBuckets.name -join "</CHOICE><CHOICE>"
                $fieldXml = '<Field DisplayName=\"Bucket\" FillInChoice=\"FALSE\" Format=\"Dropdown\" IsModern=\"TRUE\" Name=\"Bucket\" Title=\"Bucket\" Type=\"Choice\" ID=\"{e5a87c1b-14fe-4a0d-b2a1-69ed51aefe0a}\" SourceID=\"{be343550-aa79-4954-95f7-e7f3e1158888}\" StaticName=\"Bucket\" ColName=\"nvarchar7\" RowOrdinal=\"0\" Version=\"5\"><CHOICES><CHOICE>' + $bucketOptions + '</CHOICE></CHOICES></Field>';
                $field = m365 spo field add --webUrl $SiteUrl --listTitle $list.title --xml $fieldXml --options AddToAllContentTypes
                $field = m365 spo field add --webUrl $SiteUrl --listTitle $list.title --xml '<Field DisplayName=\"Assigned to\" Format=\"Dropdown\" IsModern=\"TRUE\" List=\"UserInfo\" Mult=\"TRUE\" Name=\"AssignedTo\" Title=\"AssignedTo\" Type=\"UserMulti\" UserSelectionMode=\"0\" UserSelectionScope=\"0\" ID=\"{38a2a5a8-5518-4242-9b9b-760777f5e7ea}\" SourceID=\"{50756447-36ab-447e-9b30-859b91aba49d}\" StaticName=\"AssignedTo\" />' --options AddToAllContentTypes
                $field = m365 spo field add --webUrl $SiteUrl --listTitle $list.title --xml '<Field DisplayName=\"Start date\" FriendlyDisplayFormat=\"Disabled\" Format=\"DateOnly\" IsModern=\"TRUE\" Name=\"Startdate\" Title=\"Start date\" Type=\"DateTime\" ID=\"{91ff7e76-118b-49e1-85dd-afda84856e96}\" SourceID=\"{50756447-36ab-447e-9b30-859b91aba49d}\" StaticName=\"Startdate\" ColName=\"datetime1\" RowOrdinal=\"0\" />' --options AddToAllContentTypes
                $field = m365 spo field add --webUrl $SiteUrl --listTitle $list.title --xml '<Field DisplayName=\"Due date\" FriendlyDisplayFormat=\"Disabled\" Format=\"DateOnly\" IsModern=\"TRUE\" Name=\"Duedate\" Title=\"Due date\" Type=\"DateTime\" ID=\"{23f7955a-1b1b-46f9-8b0f-316dbb37b63e}\" SourceID=\"{50756447-36ab-447e-9b30-859b91aba49d}\" StaticName=\"Duedate\" ColName=\"datetime2\" RowOrdinal=\"0\" />' --options AddToAllContentTypes
                $field = m365 spo field add --webUrl $SiteUrl --listTitle $list.title --xml '<Field AppendOnly=\"FALSE\" DisplayName=\"Checklist\" Format=\"Dropdown\" IsModern=\"TRUE\" IsolateStyles=\"TRUE\" Name=\"Checklist\" RichText=\"TRUE\" RichTextMode=\"FullHtml\" Title=\"Checklist\" Type=\"Note\" ID=\"{05f7b70d-eb7c-4428-afe6-ca9f2f22c5af}\" SourceID=\"{50756447-36ab-447e-9b30-859b91aba49d}\" StaticName=\"Checklist\" ColName=\"ntext2\" RowOrdinal=\"0\" CustomFormatter=\"\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" NumLines=\"6\" RestrictedMode=\"TRUE\" Version=\"1\" />' --options AddToAllContentTypes
                $field = m365 spo field add --webUrl $SiteUrl --listTitle $list.title --xml '<Field AppendOnly=\"FALSE\" DisplayName=\"Description\" Format=\"Dropdown\" IsModern=\"TRUE\" IsolateStyles=\"FALSE\" Name=\"Description\" RichText=\"FALSE\" RichTextMode=\"Compatible\" Title=\"Description\" Type=\"Note\" ID=\"{50d7e60e-ce7d-4428-afe6-ca9f2f22c5ce}\" SourceID=\"{50756447-36ab-447e-9b30-859b91aba49d}\" StaticName=\"Description\" ColName=\"ntext2\" RowOrdinal=\"0\" />' --options AddToAllContentTypes
                $field = m365 spo field add --webUrl $SiteUrl --listTitle $list.title --xml '<Field Type=\"Number\" DisplayName=\"Priority\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" ID=\"{a9639335-a3cf-41d2-a1fb-28d02c8ef09f}\" SourceID=\"{be343550-aa79-4954-95f7-e7f3e1158888}\" StaticName=\"Priority\" Name=\"Priority\" ColName=\"float1\" RowOrdinal=\"0\" CustomFormatter=\"{&quot;$schema&quot;:&quot;https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json&quot;,&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;flex-wrap&quot;:&quot;wrap&quot;,&quot;display&quot;:&quot;flex&quot;},&quot;children&quot;:[{&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;box-sizing&quot;:&quot;border-box&quot;,&quot;padding&quot;:&quot;4px 8px 5px 8px&quot;,&quot;display&quot;:&quot;flex&quot;,&quot;border-radius&quot;:&quot;16px&quot;,&quot;height&quot;:&quot;24px&quot;,&quot;align-items&quot;:&quot;center&quot;,&quot;white-space&quot;:&quot;nowrap&quot;,&quot;overflow&quot;:&quot;hidden&quot;,&quot;margin&quot;:&quot;4px 4px 4px 4px&quot;},&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,1]},&quot;sp-css-backgroundColor-BgRed sp-css-borderColor-WhiteFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,3]},&quot;sp-css-backgroundColor-BgPeach sp-css-borderColor-PeachFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,5]},&quot;sp-css-backgroundColor-BgGold&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,9]},&quot;&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,9]},&quot;sp-css-backgroundColor-BgMintGreen sp-css-borderColor-MintGreenFont&quot;,&quot;sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-neutralSecondary&quot;]}]}]}]}]}},&quot;children&quot;:[{&quot;elmType&quot;:&quot;span&quot;,&quot;style&quot;:{&quot;line-height&quot;:&quot;16px&quot;,&quot;height&quot;:&quot;14px&quot;},&quot;attributes&quot;:{&quot;iconName&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,1]},&quot;RingerSolid&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,3]},&quot;Important&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,5]},&quot;LocationDot&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,&quot;&quot;]},&quot;Down&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,9]},&quot;&quot;,&quot;&quot;]}]}]}]}]},&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,1]},&quot;sp-css-color-WhiteFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,3]},&quot;sp-css-color-PeachFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,5]},&quot;&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,&quot;&quot;]},&quot;&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,9]},&quot;sp-css-color-MintGreenFont&quot;,&quot;&quot;]}]}]}]}]}}},{&quot;elmType&quot;:&quot;span&quot;,&quot;style&quot;:{&quot;overflow&quot;:&quot;hidden&quot;,&quot;text-overflow&quot;:&quot;ellipsis&quot;,&quot;padding&quot;:&quot;0 3px&quot;},&quot;txtContent&quot;:&quot;=if(@currentField == 1, ''Urgent'', if(@currentField == 3, ''Important'', if(@currentField == 5, ''Medium'', if(@currentField == 9, '''', ''Low''))))&quot;,&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,1]},&quot;sp-css-color-WhiteFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,3]},&quot;sp-css-color-PeachFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,5]},&quot;sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,0]},9,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$Priority]&quot;,9]},&quot;sp-css-color-MintGreenFont&quot;,&quot;&quot;]}]}]}]}]}}}]}]}\" Version=\"23\" />'  --options AddToAllContentTypes
                $field = m365 spo field add --webUrl $SiteUrl --listTitle $list.title --xml '<Field CommaSeparator=\"TRUE\" CustomUnitOnRight=\"TRUE\" DisplayName=\"Progress\" Format=\"Dropdown\" IsModern=\"TRUE\" Name=\"Progress\" Percentage=\"FALSE\" Title=\"Progress\" Type=\"Number\" Unit=\"None\" ID=\"{abea0d9d-83b8-4d57-9c68-8b56ad4066c9}\" SourceID=\"{be343550-aa79-4954-95f7-e7f3e1158888}\" StaticName=\"Progress\" ColName=\"float2\" RowOrdinal=\"0\" CustomFormatter=\"{&quot;$schema&quot;:&quot;https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json&quot;,&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;flex-wrap&quot;:&quot;wrap&quot;,&quot;display&quot;:&quot;flex&quot;},&quot;children&quot;:[{&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;box-sizing&quot;:&quot;border-box&quot;,&quot;padding&quot;:&quot;4px 8px 5px 8px&quot;,&quot;display&quot;:&quot;flex&quot;,&quot;border-radius&quot;:&quot;16px&quot;,&quot;height&quot;:&quot;24px&quot;,&quot;align-items&quot;:&quot;center&quot;,&quot;white-space&quot;:&quot;nowrap&quot;,&quot;overflow&quot;:&quot;hidden&quot;,&quot;margin&quot;:&quot;4px 4px 4px 4px&quot;},&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,0]},&quot;sp-css-backgroundColor-BgLightGray sp-css-borderColor-LightGrayFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,50]},&quot;sp-css-backgroundColor-BgGold sp-css-borderColor-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,100]},&quot;sp-css-backgroundColor-BgGreen sp-css-borderColor-WhiteFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;&quot;]},&quot;&quot;,&quot;sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-neutralSecondary&quot;]}]}]}]}},&quot;children&quot;:[{&quot;elmType&quot;:&quot;span&quot;,&quot;style&quot;:{&quot;line-height&quot;:&quot;16px&quot;,&quot;height&quot;:&quot;14px&quot;},&quot;attributes&quot;:{&quot;iconName&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,0]},&quot;CircleRing&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,50]},&quot;CircleHalfFull&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,100]},&quot;CircleFill&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;&quot;]},&quot;&quot;,&quot;&quot;]}]}]}]},&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,0]},&quot;sp-css-color-LightGrayFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,50]},&quot;sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,100]},&quot;sp-css-color-WhiteFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;&quot;]},&quot;&quot;,&quot;&quot;]}]}]}]}}},{&quot;elmType&quot;:&quot;span&quot;,&quot;style&quot;:{&quot;overflow&quot;:&quot;hidden&quot;,&quot;text-overflow&quot;:&quot;ellipsis&quot;,&quot;padding&quot;:&quot;0 3px&quot;},&quot;txtContent&quot;:&quot;=if(@currentField == 100, ''Completed'', if(@currentField == 50, ''Completed'', if(@currentField == 0, ''Not started'', '''')))&quot;,&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,0]},&quot;sp-field-fontSizeSmall sp-css-color-LightGrayFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,50]},&quot;sp-field-fontSizeSmall sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,100]},&quot;sp-field-fontSizeSmall sp-css-color-WhiteFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;&quot;]},&quot;&quot;,&quot;&quot;]}]}]}]}}}]}]}\" Version=\"2\" />' --options AddToAllContentTypes

                $viewName = "All Items"
                $viewFields = @("Progress", "Priority", "Assigned_x0020_to", "Due date", "Start date");
                foreach ($field in $viewFields) {
                    m365 spo list view field add --webUrl $SiteUrl --listTitle  $list.title  --viewTitle $viewName --fieldTitle $field
                }
                $view = m365 spo list view set --webUrl $SiteUrl --listTitle $list.title --viewTitle $viewName --ViewQuery '<GroupBy Collapse=\"TRUE\" GroupLimit=\"30\"><FieldRef Name=\"Bucket\" /></GroupBy><OrderBy><FieldRef Name=\"ID\" /></OrderBy>'
                $view = m365 spo list view set --webUrl $SiteUrl --listTitle $list.title --viewTitle $viewName --ViewType2 "TILES"
            }

            Write-Host "Migrating tasks"

            foreach ($migrationTask in $migrationTasks) {
                $i++
                if ($ShowProgress) { Write-Host "Processing ($i/$($migrationTasks.length))" }

                $newItem = m365 spo listitem add --webUrl $SiteUrl --listTitle $plan.title --Title $migrationTask.Title --Bucket $migrationTask.Bucket --Progress $migrationTask.Progress --Description $migrationTask.Description --Assigned_x0020_to $migrationTask.AssignedTo --Priority $migrationTask.Priority | ConvertFrom-Json -AsHashtable

                # Fails with an empty date so extra check to prevent issues
                if ($migrationTask.StartDate) {
                    $updatedItem = m365 spo listitem set --webUrl $SiteUrl --listTitle $plan.title --id $newItem.Id --Start_x0020_date $migrationTask.StartDate
                }
                if ($migrationTask.DueDate) {
                    $updatedItem = m365 spo listitem set --webUrl $SiteUrl --listTitle $plan.title --id $newItem.Id --Due_x0020_date $migrationTask.DueDate
                }
                if ($migrationTask.Checklist) {
                    $updatedItem = m365 spo listitem set --webUrl $SiteUrl --listTitle $plan.title --id $newItem.Id --Checklist $migrationTask.Checklist
                }
            }
            
            Write-Host "Migrating finished for the plan $($plan.title)"
        } else {
            Write-Host "Skipping migrating due to lack of tasks"
        }
    }
    ```

Keywords:

- CLI for Microsoft 365
- SharePoint Online
- Planner
