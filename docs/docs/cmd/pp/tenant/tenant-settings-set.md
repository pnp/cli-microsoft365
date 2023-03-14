# pp tenant settings set

Sets the global Power Platform configuration of the tenant

## Usage

```sh
m365 pp tenant settings set [options]
```

## Options

`--walkMeOptOut [walkMeOptOut]`
: Ability to opt out of guided experiences using WalkMe in Power Platform. Valid values: `true`, `false`.

`--disableNPSCommentsReachout [disableNPSCommentsReachout]`
: Ability to disable re-surveying users who left prior feedback via NPS prompts in Power Platform. Valid values: `true`, `false`.

`--disableNewsletterSendout [disableNewsletterSendout]`
: Ability to disable the newsletter sendout feature. Valid values: `true`, `false`.

`--disableEnvironmentCreationByNonAdminusers [disableEnvironmentCreationByNonAdminusers]`
: Restrict all environments to be created by Global Admins, Power Platform Admins, or Dynamics365 Service Admins. Valid values: `true`, `false`.

`--disablePortalsCreationByNonAdminusers [disablePortalsCreationByNonAdminusers]`
: Restrict all portals to be created by Global Admins, Power Platform Admins, or Dynamics365 Service Admins. Valid values: `true`, `false`.

`--disableSurveyFeedback [disableSurveyFeedback]`
: Ability to disable all NPS survey feedback prompts in Power Platform. Valid values: `true`, `false`.

`--disableTrialEnvironmentCreationByNonAdminusers [disableTrialEnvironmentCreationByNonAdminusers]`
: Restrict all trial environments to be created by Global Admins, Power Platform Admins, or Dynamics365 Service Admins. Valid values: `true`, `false`.

`--disableCapacityAllocationByEnvironmentAdmins [disableCapacityAllocationByEnvironmentAdmins]`
: Ability to disable capacity allocation by environment administrators. Valid values: `true`, `false`.

`--disableSupportTicketsVisibleByAllUsers [disableSupportTicketsVisibleByAllUsers]`
: Ability to disable support ticket creation by non-admin users in the tenant. Valid values: `true`, `false`.

`--disableDocsSearch [disableDocsSearch]`
: When this setting is true, users in the environment will see a message that Microsoft Learn and Documentation search categories have been turned off by the administrator in the search results page. Valid values: `true`, `false`.

`--disableCommunitySearch [disableCommunitySearch]`
: When this setting is true, users in the environment will see a message that Community and Blog search categories have been turned off by the administrator in the search results page. Valid values: `true`, `false`.

`--disableBingVideoSearch [disableBingVideoSearch]`
: When this setting is true, users in the environment will see a message that Video search categories have been turned off by the administrator in the search results page. Valid values: `true`, `false`.

`--shareWithColleaguesUserLimit [shareWithColleaguesUserLimit]`
: Maximum value setting for the number of users in a security group used to share an app built using Power Apps on Microsoft Teams. Specify any number as a value.

`--disableShareWithEveryone [disableShareWithEveryone]`
: Ability to disable the Share With Everyone capability in all Power Apps. Valid values: `true`, `false`.

`--enableGuestsToMake [enableGuestsToMake]`
: Ability to allow guest users in your tenant to create Power Apps. Valid values: `true`, `false`.

`--disableAdminDigest [disableAdminDigest]`
: When true, the entire organization is unsubscribed from the weekly digest. Valid values: `true`, `false`.

`--disableDeveloperEnvironmentCreationByNonAdminUsers [disableDeveloperEnvironmentCreationByNonAdminUsers]`
: Restrict all developer environments to be created by Global Admins, Power Platform Admins, or Dynamics365 Service Admins. Valid values: `true`, `false`.

`--disableBillingPolicyCreationByNonAdminUsers [disableBillingPolicyCreationByNonAdminUsers]`
: Restrict billing policies to be created by Global Admins, Power Platform Admins, or Dynamics365 Service Admins. Valid values: `true`, `false`.

`--disableChampionsInvitationReachout [disableChampionsInvitationReachout]`
: Ability to disable all invitations to become a Power Platform champion. Valid values: `true`, `false`.

`--disableSkillsMatchInvitationReachout [disableSkillsMatchInvitationReachout]`
: Ability to disable all skills match invitations to become part of the makers community. Valid values: `true`, `false`.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Disable environment creation for non-admin users

```sh
m365 pp tenant settings set --disableEnvironmentCreationByNonAdminUsers true --disableTrialEnvironmentCreationByNonAdminUsers true --disableDeveloperEnvironmentCreationByNonAdminUsers true
```

Enable Power App creation for guest users

```sh
m365 pp tenant settings set --enableGuestsToMake true
```

Disable guided experience, survey feedback and newsletter

```sh
m365 pp tenant settings set --walkMeOptOut true --disableNewsletterSendout true --disableSurveyFeedback true
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
    # pp tenant settings set --disableEnvironmentCreationByNonAdminUsers false --disableTrialEnvironmentCreationByNonAdminUsers false --disableDeveloperEnvironmentCreationByNonAdminUsers false

    Date: 14/3/2023

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
