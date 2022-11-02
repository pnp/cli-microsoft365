# pa app list

Lists all Power Apps apps

## Usage

```sh
pa app list [options]
```

## Options

`-e, --environment [environment]`
: The name of the environment for which to retrieve available apps

`--asAdmin`
: Set, to list all Power Apps as admin. Otherwise will return only your own apps

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reaches general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

By default, the `app list` command returns only your apps. To list all apps, use the `asAdmin` option and make sure to specify the `environment` option. You cannot specify only one of the options, when specifying the `environment` option the `asAdmin` option has to be present as well.

## Examples

List all your apps

```sh
m365 pa app list
```

List all apps in a given environment

```sh
m365 pa app list --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --asAdmin
```

## Response

=== "JSON"

    ```json
    [
      {
        "name":"37ea6004-f07b-46ca-8ef3-a256b67b4dbb",
        "id":"/providers/Microsoft.PowerApps/apps/37ea6004-f07b-46ca-8ef3-a256b67b4dbb",
        "type":"Microsoft.PowerApps/apps",
        "tags":{
          "primaryDeviceWidth":"1366",
          "primaryDeviceHeight":"768",
          "supportsPortrait":"true",
          "supportsLandscape":"true",
          "primaryFormFactor":"Tablet",
          "publisherVersion":"3.22102.32",
          "minimumRequiredApiVersion":"2.2.0",
          "hasComponent":"false",
          "hasUnlockedComponent":"false",
          "isUnifiedRootApp":"false",
          "sienaVersion":"20221025T212812Z-3.22102.32.0"
        },
        "properties":{
          "appVersion":"2022-10-25T21:28:12Z",
          "lastDraftVersion":"2022-10-25T21:28:12Z",
          "lifeCycleId":"Published",
          "status":"Ready",
          "createdByClientVersion":{
            "major":3,
            "minor":22102,
            "build":32,
            "revision":0,
            "majorRevision":0,
            "minorRevision":0
          },
          "minClientVersion":{
            "major":3,
            "minor":22102,
            "build":32,
            "revision":0,
            "majorRevision":0,
            "minorRevision":0
          },
          "owner":{
            "id":"fe36f75e-c103-410b-a18a-2bf6df06ac3a",
            "displayName":"contoso",
            "email":"user@contoso.onmicrosoft.com",
            "type":"User",
            "tenantId":"e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
            "userPrincipalName":"user@contoso.onmicrosoft.com"
          },
          "createdBy":{
            "id":"fe36f75e-c103-410b-a18a-2bf6df06ac3a",
            "displayName":"contoso",
            "email":"user@contoso.onmicrosoft.com",
            "type":"User",
            "tenantId":"e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
            "userPrincipalName":"user@contoso.onmicrosoft.com"
          },
          "lastModifiedBy":{
            "id":"fe36f75e-c103-410b-a18a-2bf6df06ac3a",
            "displayName":"contoso",
            "email":"user@contoso.onmicrosoft.com",
            "type":"User",
            "tenantId":"e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
            "userPrincipalName":"user@contoso.onmicrosoft.com"
          },
          "lastPublishedBy":{
            "id":"fe36f75e-c103-410b-a18a-2bf6df06ac3a",
            "displayName":"contoso",
            "email":"user@contoso.onmicrosoft.com",
            "type":"User",
            "tenantId":"e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
            "userPrincipalName":"user@contoso.onmicrosoft.com"
          },
          "backgroundColor":"RGBA(0,176,240,1)",
          "backgroundImageUri":"https://pafeblobprodam.blob.core.windows.net:443/20221025t000000zddd642012aba4021a4886c8e21a3e1cb/logoSmallFile?sv=2018-03-28&sr=c&sig=cOkbwChyhCO%2BEJpqMDRxrXaxRoPD1TbTy%2B%2BFkdJEOjI%3D&se=2022-12-24T10%3A06%3A27Z&sp=rl",
          "teamsColorIconUrl":"https://pafeblobprodam.blob.core.windows.net:443/20221025t000000ze297221f3dc643ed9686b72b22d9a414/teamscoloricon.png?sv=2018-03-28&sr=c&sig=Fhk8E0LO4Lw0mHvNawCF5Ld7GHzPHo9l7RxvZbeZY48%3D&se=2022-12-24T10%3A06%3A27Z&sp=rl",
          "teamsOutlineIconUrl":"https://pafeblobprodam.blob.core.windows.net:443/20221025t000000ze297221f3dc643ed9686b72b22d9a414/teamsoutlineicon.png?sv=2018-03-28&sr=c&sig=Fhk8E0LO4Lw0mHvNawCF5Ld7GHzPHo9l7RxvZbeZY48%3D&se=2022-12-24T10%3A06%3A27Z&sp=rl",
        "displayName":"Test application",
          "description":"",
          "commitMessage":"",
          "appUris":{
            "documentUri":{
              "value":"https://pafeblobprodam.blob.core.windows.net:443/20221025t000000zddd642012aba4021a4886c8e21a3e1cb/document.msapp?sv=2018-03-28&sr=c&sig=laSGdpZL03POyAABXvdsyxv8YDDB8JPZIBccztIe39Q%3D&se=2022-11-04T12%3A00%3A00Z&sp=rl",
              "readonlyValue":"https://pafeblobprodam-secondary.blob.core.windows.net/20221025t000000zddd642012aba4021a4886c8e21a3e1cb/document.msapp?sv=2018-03-28&sr=c&sig=laSGdpZL03POyAABXvdsyxv8YDDB8JPZIBccztIe39Q%3D&se=2022-11-04T12%3A00%3A00Z&sp=rl"
            },
            "imageUris":[],
            "additionalUris":[]
          },
          "createdTime":"2022-10-25T21:28:12.7171469Z",
          "lastModifiedTime":"2022-10-25T21:28:12.7456297Z",
          "lastPublishTime":"2022-10-25T21:28:12Z",
          "sharedGroupsCount":0,
          "sharedUsersCount":0,
          "appOpenProtocolUri":"ms-apps:///providers/Microsoft.PowerApps/apps/37ea6004-f07b-46ca-8ef3-a256b67b4dbb",
          "appOpenUri":"https://apps.powerapps.com/play/e/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/a/37ea6004-f07b-46ca-8ef3-a256b67b4dbb?tenantId=e1dd4023-a656-480a-8a0e-c1b1eec51e1d&hint=296b0ef7-b4d0-4124-b835-f9c220a1f4bc",
          "appPlayUri":"https://apps.powerapps.com/play/e/default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/a/37ea6004-f07b-46ca-8ef3-a256b67b4dbb?tenantId=e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
          "appPlayEmbeddedUri":"https://apps.powerapps.com/play/e/default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/a/37ea6004-f07b-46ca-8ef3-a256b67b4dbb?tenantId=e1dd4023-a656-480a-8a0e-c1b1eec51e1d&hint=296b0ef7-b4d0-4124-b835-f9c220a1f4bc&telemetryLocation=eu",
          "appPlayTeamsUri":"https://apps.powerapps.com/play/e/default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/a/37ea6004-f07b-46ca-8ef3-a256b67b4dbb?tenantId=e1dd4023-a656-480a-8a0e-c1b1eec51e1d&source=teamstab&hint=296b0ef7-b4d0-4124-b835-f9c220a1f4bc&telemetryLocation=eu&locale={locale}&channelId={channelId}&channelType={channelType}&chatId={chatId}&groupId={groupId}&hostClientType={hostClientType}&isFullScreen={isFullScreen}&entityId={entityId}&subEntityId={subEntityId}&teamId={teamId}&teamType={teamType}&theme={theme}&userTeamRole={userTeamRole}",
          "userAppMetadata":{
            "favorite":"NotSpecified",
            "includeInAppsList":true
          },
          "isFeaturedApp":false,
          "bypassConsent":false,
          "isHeroApp":false,
          "environment":{
            "id":"/providers/Microsoft.PowerApps/environments/default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
            "name":"default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d"
          },
          "appPackageDetails":{
            "playerPackage":{
              "value":"https://pafeblobprodam.blob.core.windows.net:443/20221025t000000zac5237a2672a4782ad5a7d71040c032b/5b38cd68-a930-4a14-be71-c622de887d1a/player.msappk?sv=2018-03-28&sr=c&sig=eztEkTd1pFaFEITA%2Bqqj2XCpxwgeujMC7FahMmEkujA%3D&se=2022-11-04T12%3A00%3A00Z&sp=rl",
              "readonlyValue":"https://pafeblobprodam-secondary.blob.core.windows.net/20221025t000000zac5237a2672a4782ad5a7d71040c032b/5b38cd68-a930-4a14-be71-c622de887d1a/player.msappk?sv=2018-03-28&sr=c&sig=eztEkTd1pFaFEITA%2Bqqj2XCpxwgeujMC7FahMmEkujA%3D&se=2022-11-04T12%3A00%3A00Z&sp=rl"
            },
            "webPackage":{
              "value":"https://pafeblobprodam.blob.core.windows.net:443/20221025t000000zac5237a2672a4782ad5a7d71040c032b/5b38cd68-a930-4a14-be71-c622de887d1a/web/index.web.html?sv=2018-03-28&sr=c&sig=eztEkTd1pFaFEITA%2Bqqj2XCpxwgeujMC7FahMmEkujA%3D&se=2022-11-04T12%3A00%3A00Z&sp=rl",
              "readonlyValue":"https://pafeblobprodam-secondary.blob.core.windows.net/20221025t000000zac5237a2672a4782ad5a7d71040c032b/5b38cd68-a930-4a14-be71-c622de887d1a/web/index.web.html?sv=2018-03-28&sr=c&sig=eztEkTd1pFaFEITA%2Bqqj2XCpxwgeujMC7FahMmEkujA%3D&se=2022-11-04T12%3A00%3A00Z&sp=rl"
            },
            "unauthenticatedWebPackage":{
              "value":"https://pafeblobprodam.blob.core.windows.net/alt20221025t000000z529d41282a634bf6b94383dde5a8d52c/20221025T212824Z/index.web.html"
            },
            "documentServerVersion":{
              "major":3,
              "minor":22102,
              "build":33,
              "revision":0,
              "majorRevision":0,
              "minorRevision":0
            },
            "appPackageResourcesKind":"Split",
            "packagePropertiesJson":"{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/kdfj31mdao7t9/preloadindex.web.html\",\"id\":\"638023301009567627\",\"v\":2.1}",
            "id":"20221025t000000zac5237a2672a4782ad5a7d71040c032bhttps://pafeblobprodam.blob.core.windows.net/20221025t000000zac5237a2672a4782ad5a7d71040c032b/5b38cd68-a930-4a14-be71-c622de887d1a/web/index.web.html?sv=2018-03-28&sr=c&sig=eztEkTd1pFaFEITA%2Bqqj2XCpxwgeujMC7FahMmEkujA%3D&se=2022-11-04T12%3A00%3A00Z&sp=rlhttps://pafeblobprodam.blob.core.windows.net/20221025t000000zac5237a2672a4782ad5a7d71040c032b/5b38cd68-a930-4a14-be71-c622de887d1a/player.msappk?sv=2018-03-28&sr=c&sig=eztEkTd1pFaFEITA%2Bqqj2XCpxwgeujMC7FahMmEkujA%3D&se=2022-11-04T12%3A00%3A00Z&sp=rlhttps://pafeblobprodam.blob.core.windows.net/alt20221025t000000z529d41282a634bf6b94383dde5a8d52c/20221025T212824Z/index.web.html"
          },
          "almMode":"Environment",
          "performanceOptimizationEnabled":true,
          "unauthenticatedWebPackageHint":"296b0ef7-b4d0-4124-b835-f9c220a1f4bc",
          "canConsumeAppPass":true,
          "enableModernRuntimeMode":false,
          "executionRestrictions":{
            "isTeamsOnly":false,
            "dataLossPreventionEvaluationResult":{
              "status":"Compliant",
              "lastEvaluationDate":"2022-10-25T21:28:30.2281817Z",
              "violations":[],
              "violationsByPolicy":[],
              "violationErrorMessage":"De app gebruikt de volgende connectors: ."
            }
          },
          "appPlanClassification":"Standard",
          "usesPremiumApi":false,
          "usesOnlyGrandfatheredPremiumApis":true,
          "usesCustomApi":false,
          "usesOnPremiseGateway":false,
          "usesPcfExternalServiceUsage":false,
          "isCustomizable":true
        },
        "appLocation":"europe",
        "appType":"ClassicCanvasApp",
        "displayName":"PowerApps Application"
      }
    ]
    ```

=== "Text"

    ```text
    name                                 displayName
    ------------------------------------ ---------------------
    37ea6004-f07b-46ca-8ef3-a256b67b4dbb PowerApps Application   
    ```

=== "CSV"

    ```csv
    name,displayName
    37ea6004-f07b-46ca-8ef3-a256b67b4dbb,"PowerApps Application"
    ```
