# purview sensitivitylabel policysettings list

Get a list of policy settings for a sensitivity label.

## Usage

```sh
m365 purview sensitivitylabel policysettings list [options]
```

## Options

`--userId [userId]`
: User's Azure AD ID. Optionally specify this if you want to get a list of policy settings for a sensitivity label that the user has access to. Specify either `userId` or `userName` but not both.

`--userName [userName]`
: User's UPN (user principal name, e.g. johndoe@example.com). Optionally specify this if you want to get a list of policy settings for a sensitivity label that the user has access to. Specify either `userId` or `userName` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Get a list of policy settings for a sensitivity label.

```sh
m365 purview sensitivitylabel policysettings list
```

Get a list of policy settings for a sensitivity label that a specific user has access to by its Id.

```sh
m365 purview sensitivitylabel policysettings list --userId 59f80e08-24b1-41f8-8586-16765fd830d3
```

Get a list of policy settings for a sensitivity label that a specific user has access to by its UPN.

```sh
m365 purview sensitivitylabel policysettings list --userName john.doe@contoso.com
```

## Remarks

!!! attention
    This command is based on a Microsoft Graph API that is currently in preview and is subject to change once the API reached general availability.

!!! attention
    When operating in app-only mode, you have the option to use either the `userName` or `userId` parameters to retrieve the sensitivity policy settings for a specific user. Without specifying either of these parameters, the command will retrieve the sensitivity policy settings for the currently authenticated user when operating in delegated mode.


## Response

=== "JSON"

    ```json
    {
      "id": "71F139249895C2F6DC861031DAC47E0C2C37C6595582D4248CC77FD7293681B5DE348BC71AEB44068CB397DB021CADB4",
      "moreInfoUrl": "https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-with-sensitivity-labels?view=o365-worldwide#end-user-documentation-for-sensitivity-labels",
      "isMandatory": true,
      "isDowngradeJustificationRequired": true,
      "defaultLabelId": "022bb90d-0cda-491d-b861-d195b14532dc"
    }
    ```

=== "Text"

    ```text
    defaultLabelId                  : 022bb90d-0cda-491d-b861-d195b14532dc
    id                              : 71F139249895C2F6DC861031DAC47E0C2C37C6595582D4248CC77FD7293681B5DE348BC71AEB44068CB397
    DB021CADB4
    isDowngradeJustificationRequired: true
    isMandatory                     : true
    moreInfoUrl                     : https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-with-sensitivity-labels?view=o365-worldwide#end-user-documentation-for-sensitivity-labels
    ```

=== "CSV"

    ```csv
    id,moreInfoUrl,isMandatory,isDowngradeJustificationRequired,defaultLabelId
    71F139249895C2F6DC861031DAC47E0C2C37C6595582D4248CC77FD7293681B5DE348BC71AEB44068CB397DB021CADB4,https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-with-sensitivity-labels?view=o365-worldwide#end-user-documentation-for-sensitivity-labels,1,1,022bb90d-0cda-491d-b861-d195b14532dc
    ```

=== "Markdown"

    ```md
    # purview sensitivitylabel policysettings list

    Date: 4/11/2023

    ## 71F139249895C2F6DC861031DAC47E0C2C37C6595582D4248CC77FD7293681B5DE348BC71AEB44068CB397DB021CADB4

    Property | Value
    ---------|-------
    id | 71F139249895C2F6DC861031DAC47E0C2C37C6595582D4248CC77FD7293681B5DE348BC71AEB44068CB397DB021CADB4
    moreInfoUrl | https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-with-sensitivity-labels?view=o365-worldwide#end-user-documentation-for-sensitivity-labels
    isMandatory | true
    isDowngradeJustificationRequired | true
    defaultLabelId | 022bb90d-0cda-491d-b861-d195b14532dc
    ```
