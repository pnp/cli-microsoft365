# Add multiple tasks in Planner

Author: [Veronique Lengelle](https://veronicageek.com/2019/get-files-with-specific-names/)

## Add multiple tasks using a CSV file

This script will create multiple tasks to a Planner plan from the information provided in your `csv` file. For this particular example, the `csv` file should contain the following columns: `TaskName`, `Description`, `PlanID`, `BucketName`, `StartDateTime`, `AssignedToUserNames`.

=== "PowerShell"

    ```powershell
    $m365Status = m365 status --output text
    if ($m365Status -eq "Logged Out") {
      # Connection to Microsoft 365
      m365 login
    }

    #Import your CSV file
    $csvTasks = Import-Csv -Path "<YOUR-CSV-PATH>"

    foreach($task in $csvTasks){
      m365 planner task add --title "$($task.TaskName)" --description "$($task.Description)" --planId "$($task.PlanID)" --bucketName "$($task.BucketName)" --startDateTime "$($task.StartDateTime)" --assignedToUserNames "$($task.AssignedToUserNames)"
    }
    ```

## Create multiple tasks using an in-script hashtable

=== "PowerShell"

    ```powershell
    #Create multiple tasks onto a single Planner bucket with in-script tasks
    $allMyTasks = @{
      Taks1 = @{
        TaskName = "Task 1"
        Description = "Description 1"
        PlanID = "fdmtzs0rkkik0ILStJRu12345678"
        BucketName = "SharePoint"
        StartDateTime = "2022-06-01T09:30:00.000Z"
        DueDateTime = "2022-06-02T17:30:00.000Z"
        AssignedToUserNames = "veronique@contoso.onmicrosoft.com"
      }
      Task2 = @{
        TaskName = "Task 2"
        Description = "Description 2"
        PlanID = "fdmtzs0rkkik0ILStJRu12345678"
        BucketName = "PowerApps"
        StartDateTime = "2022-06-05T09:30:00.000Z"
        DueDateTime = "2022-06-07T17:30:00.000Z"
        AssignedToUserNames = "veronique@contoso.onmicrosoft.com, jdoe@contoso.onmicrosoft.com"
      }
    }

    foreach($task in $allMyTasks.Values){
      m365 planner task add --title "$($task.TaskName)" --description "$($task.Description)" --planId "$($task.PlanID)" --bucketName "$($task.BucketName)" --startDateTime "$($task.StartDateTime)" --dueDateTime "$($task.DueDateTime)" --assignedToUserNames "$($task.AssignedToUserNames)"
    }
    ```

Keywords

-   Planner
