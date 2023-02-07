# pp tenant settings list

Lists the global Power Platform tenant settings

## Usage

```sh
m365 pp tenant settings list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Lists the global Power Platform settings of the tenant

```sh
m365 pp tenant settings list
```

## Response

=== "JSON"

    ```json
    {
      "walkMeOptOut": false,
      "disableNPSCommentsReachout": false,
      "disableNewsletterSendout": false,
      "disableEnvironmentCreationByNonAdminUsers": false,
      "disablePortalsCreationByNonAdminUsers": false,
      "disableSurveyFeedback": false,
      "disableTrialEnvironmentCreationByNonAdminUsers": false,
      "disableCapacityAllocationByEnvironmentAdmins": false,
      "disableSupportTicketsVisibleByAllUsers": false,
      "powerPlatform": {
        "search": {
          "disableDocsSearch": false,
          "disableCommunitySearch": false,
          "disableBingVideoSearch": false
        },
        "teamsIntegration": {
          "shareWithColleaguesUserLimit": 10000
        },
        "powerApps": {
          "disableShareWithEveryone": false,
          "enableGuestsToMake": false,
          "disableMembersIndicator": false
        },
        "environments": {},
        "governance": {
          "disableAdminDigest": false,
          "allowDeveloperEnvironmentProvisioning": false
        },
        "licensing": {
          "disableBillingPolicyCreationByNonAdminUsers": false
        },
        "powerPages": {}
      }
    }
    ```

=== "Text"

    ```text
    disableCapacityAllocationByEnvironmentAdmins  : false
    disableEnvironmentCreationByNonAdminUsers     : false
    disableNPSCommentsReachout                    : false
    disablePortalsCreationByNonAdminUsers         : false
    disableSupportTicketsVisibleByAllUsers        : false
    disableSurveyFeedback                         : false
    disableTrialEnvironmentCreationByNonAdminUsers: false
    walkMeOptOut                                  : false
    ```

=== "CSV"

    ```csv
    disableCapacityAllocationByEnvironmentAdmins,disableEnvironmentCreationByNonAdminUsers,disableNPSCommentsReachout,disablePortalsCreationByNonAdminUsers,disableSupportTicketsVisibleByAllUsers,disableSurveyFeedback,disableTrialEnvironmentCreationByNonAdminUsers,walkMeOptOut
    false,false,false,false,false,false,false,false
    ```

=== "Markdown"

    ```md
    # pp tenant settings list

    Date: 6/2/2023

    Property | Value
    ---------|-------
    walkMeOptOut | false
    disableNPSCommentsReachout | false
    disableNewsletterSendout | false
    disableEnvironmentCreationByNonAdminUsers | false
    disablePortalsCreationByNonAdminUsers | false
    disableSurveyFeedback | false
    disableTrialEnvironmentCreationByNonAdminUsers | false
    disableCapacityAllocationByEnvironmentAdmins | false
    disableSupportTicketsVisibleByAllUsers | false
    powerPlatform | {"search":{"disableDocsSearch":false,"disableCommunitySearch":false,"disableBingVideoSearch":false},"teamsIntegration":{"shareWithColleaguesUserLimit":10000},"powerApps":{"disableShareWithEveryone":false,"enableGuestsToMake":false,"disableMembersIndicator":false},"environments":{},"governance":{"disableAdminDigest":false,"disableDeveloperEnvironmentCreationByNonAdminUsers":false},"licensing":{"disableBillingPolicyCreationByNonAdminUsers":false},"powerPages":{}}
    ```
