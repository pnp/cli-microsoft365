import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';
const command: Command = require('./app-owner-set');

describe(commands.APP_OWNER_SET, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const validEnvironmentName: string = 'Default-6a2903af-9c03-4c02-a50b-e7419599925b';
  const validAppName: string = '784670e6-199a-4993-ae13-4b6747a0cd5d';
  const validUserId: string = 'd2481133-e3ed-4add-836d-6e200969dd03';
  const validUserName: string = 'john.doe@contoso.com';
  const userResponse = { value: [{ id: validUserId }] };
  const appOwnerSetResponse: any = {
    name: "5806ec64-b473-4e69-9dd4-3cf606c8c518",
    id: "/providers/Microsoft.PowerApps/apps/5806ec64-b473-4e69-9dd4-3cf606c8c518",
    type: "Microsoft.PowerApps/apps",
    tags: {
      primaryDeviceWidth: "640",
      primaryDeviceHeight: "1136",
      supportsPortrait: "true",
      supportsLandscape: "true",
      primaryFormFactor: "Phone",
      publisherVersion: "3.22111.19",
      minimumRequiredApiVersion: "2.2.0",
      hasComponent: false,
      hasUnlockedComponent: false,
      isUnifiedRootApp: false,
      sienaVersion: "20221129T133625Z-3.22111.19.0"
    },
    properties: {
      appVersion: "2022-11-29T13:36:25Z",
      lastDraftVersion: "2022-11-29T13:36:25Z",
      lifeCycleId: "Published",
      status: "Ready",
      createdByClientVersion: {
        major: 3,
        minor: 22111,
        build: 19,
        revision: 0,
        majorRevision: 0,
        minorRevision: 0
      },
      minClientVersion: {
        major: 3,
        minor: 22111,
        build: 19,
        revision: 0,
        majorRevision: 0,
        minorRevision: 0
      },
      owner: {
        id: "643e3489-67b3-49f7-ab80-2ebc6402e5a0",
        displayName: "John Doe",
        email: "john.doe@contoso.com",
        type: "User",
        tenantId: "3f31450a-fa3c-4232-beed-a599786feee7",
        userPrincipalName: "john.doe@contoso.com"
      },
      createdBy: {
        id: "643e3489-67b3-49f7-ab80-2ebc6402e5a0",
        displayName: "John Doe",
        email: "john.doe@contoso.com",
        type: "User",
        tenantId: "3f31450a-fa3c-4232-beed-a599786feee7",
        userPrincipalName: "john.doe@contoso.com"
      },
      lastModifiedBy: {
        id: "643e3489-67b3-49f7-ab80-2ebc6402e5a0",
        displayName: "John Doe",
        email: "john.doe@contoso.com",
        type: "User",
        tenantId: "3f31450a-fa3c-4232-beed-a599786feee7",
        userPrincipalName: "john.doe@contoso.com"
      },
      lastPublishedBy: {
        id: "643e3489-67b3-49f7-ab80-2ebc6402e5a0",
        displayName: "John Doe",
        email: "john.doe@contoso.com",
        type: "User",
        tenantId: "3f31450a-fa3c-4232-beed-a599786feee7",
        userPrincipalName: "john.doe@contoso.com"
      },
      backgroundColor: "RGBA(0,176,240,1)",
      backgroundImageUri: "https://pafeblobprodam.blob.core.windows.net:443/20221129t000000z94276e1625484bbcb5dba13f553b209a/logoSmallFile?sv=2018-03-28&sr=c&sig=cA2JCb6eYE%2B%2Fyrscllm4LpzDW%2FIl6wNUbGr6vMUawUQ%3D&se=2023-08-21T13%3A23%3A50Z&sp=rl",
      teamsColorIconUrl: "https://pafeblobprodam.blob.core.windows.net:443/20221129t000000z277f90efa7134656b26a790470ecb1e8/teamscoloricon.png?sv=2018-03-28&sr=c&sig=4GfzXwhmgkI8QqwPuvhAwM72Qd4Mv4wfL%2Fq4mOwEOwo%3D&se=2023-08-21T13%3A23%3A50Z&sp=rl",
      teamsOutlineIconUrl: "https://pafeblobprodam.blob.core.windows.net:443/20221129t000000z277f90efa7134656b26a790470ecb1e8/teamsoutlineicon.png?sv=2018-03-28&sr=c&sig=4GfzXwhmgkI8QqwPuvhAwM72Qd4Mv4wfL%2Fq4mOwEOwo%3D&se=2023-08-21T13%3A23%3A50Z&sp=rl",
      displayName: "App",
      description: "",
      commitMessage: "",
      appUris: {
        documentUri: {
          value: "https://pafeblobprodam.blob.core.windows.net:443/20221129t000000z94276e1625484bbcb5dba13f553b209a/document.msapp?sv=2018-03-28&sr=c&sig=Tfqtg%2Ba7WP%2FCmsYyObiwDS%2FVp%2BsPJw%2FxxCFGpRFFxXk%3D&se=2023-06-30T00%3A00%3A00Z&sp=rl",
          readonlyValue: "https://pafeblobprodam-secondary.blob.core.windows.net/20221129t000000z94276e1625484bbcb5dba13f553b209a/document.msapp?sv=2018-03-28&sr=c&sig=Tfqtg%2Ba7WP%2FCmsYyObiwDS%2FVp%2BsPJw%2FxxCFGpRFFxXk%3D&se=2023-06-30T00%3A00%3A00Z&sp=rl"
        },
        imageUris: [],
        additionalUris: []
      },
      createdTime: "2022-11-29T13:36:25.3303525Z",
      lastModifiedTime: "2023-06-23T16:26:12.4391941Z",
      lastPublishTime: "2022-11-29T13:36:25Z",
      sharedGroupsCount: 0,
      sharedUsersCount: 0,
      appOpenProtocolUri: "ms-apps:///providers/Microsoft.PowerApps/apps/5806ec64-b473-4e69-9dd4-3cf606c8c518",
      appOpenUri: "https://apps.powerapps.com/play/e/Default-3f31450a-fa3c-4232-beed-a599786feee7/a/5806ec64-b473-4e69-9dd4-3cf606c8c518?tenantId=3f31450a-fa3c-4232-beed-a599786feee7&hint=8ba91c45-c476-4d51-9f11-93e5bd3a1a35",
      appPlayUri: "https://apps.powerapps.com/play/e/default-3f31450a-fa3c-4232-beed-a599786feee7/a/5806ec64-b473-4e69-9dd4-3cf606c8c518?tenantId=3f31450a-fa3c-4232-beed-a599786feee7",
      appPlayEmbeddedUri: "https://apps.powerapps.com/play/e/default-3f31450a-fa3c-4232-beed-a599786feee7/a/5806ec64-b473-4e69-9dd4-3cf606c8c518?tenantId=3f31450a-fa3c-4232-beed-a599786feee7&hint=8ba91c45-c476-4d51-9f11-93e5bd3a1a35&telemetryLocation=eu",
      appPlayTeamsUri: "https://apps.powerapps.com/play/e/default-3f31450a-fa3c-4232-beed-a599786feee7/a/5806ec64-b473-4e69-9dd4-3cf606c8c518?tenantId=3f31450a-fa3c-4232-beed-a599786feee7&source=teamstab&hint=8ba91c45-c476-4d51-9f11-93e5bd3a1a35&telemetryLocation=eu&locale={locale}&channelId={channelId}&channelType={channelType}&chatId={chatId}&groupId={groupId}&hostClientType={hostClientType}&isFullScreen={isFullScreen}&entityId={entityId}&subEntityId={subEntityId}&teamId={teamId}&teamType={teamType}&theme={theme}&userTeamRole={userTeamRole}",
      databaseReferences: {
        'default.cds': {
          databaseDetails: {
            referenceType: "Environmental",
            environmentName: "default.cds",
            overrideValues: {
              status: "NotSpecified"
            }
          },
          dataSources: {
            '43b0dc67-68af-47ae-9a9c-9e69d396ba3e': {
              entitySetName: "msdyn_aimodels",
              logicalName: "msdyn_aimodel"
            }
          },
          actions: [
            "providers/PowerPlatform.Governance/Operations/Read",
            "providers/PowerPlatform.Governance/Operations/Read"
          ]
        }
      },
      userAppMetadata: {
        favorite: "NotSpecified",
        includeInAppsList: true
      },
      isFeaturedApp: false,
      bypassConsent: false,
      isHeroApp: false,
      environment: {
        id: "/providers/Microsoft.PowerApps/environments/default-3f31450a-fa3c-4232-beed-a599786feee7",
        name: "default-3f31450a-fa3c-4232-beed-a599786feee7"
      },
      almMode: "Environment",
      performanceOptimizationEnabled: true,
      unauthenticatedWebPackageHint: "8ba91c45-c476-4d51-9f11-93e5bd3a1a35",
      canConsumeAppPass: true,
      enableModernRuntimeMode: false,
      executionRestrictions: {
        isTeamsOnly: false,
        dataLossPreventionEvaluationResult: {
          status: "Compliant",
          lastEvaluationDate: "2022-11-29T13:36:28.7219737Z",
          violations: [],
          violationsByPolicy: [],
          violationErrorMessage: "The app uses the following connectors: shared_commondataservice."
        }
      },
      appPlanClassification: "Premium",
      usesPremiumApi: true,
      usesOnlyGrandfatheredPremiumApis: false,
      usesCustomApi: false,
      usesOnPremiseGateway: false,
      usesPcfExternalServiceUsage: false,
      isCustomizable: true,
      appDocumentComplexity: {
        controlCount: 5,
        pcfControlCount: 1,
        uiComponentsCount: 0,
        totalRuleLengthOnStart: 0,
        dataSourceCount: 5,
        totalRuleLengthHistogram: [
          7,
          9,
          13,
          1,
          0,
          0,
          0,
          0,
          0,
          0,
          0,
          0
        ],
        blocksOnStart: false,
        namedFormulasCount: 0,
        startScreenUsed: false
      }
    },
    appLocation: "europe",
    appType: "ClassicCanvasApp"
  };

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_OWNER_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId or userName not specified', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, appName: validAppName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId and userName are both specified', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, appName: validAppName, userId: validUserId, userName: validUserName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appName is not a GUID', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, appName: 'invalid', userId: validUserId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if appName is a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, appName: validAppName, userId: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userId is not a GUID', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, appName: validAppName, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, appName: validAppName, userId: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, appName: validAppName, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userName is a valid UPN', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, appName: validAppName, userName: validUserName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if roleForOldAppOwner is not a valid role', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, appName: validAppName, userId: validUserId, roleForOldAppOwner: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if roleForOldAppOwner is a valid role', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, appName: validAppName, userId: validUserId, roleForOldAppOwner: 'CanView' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('sets a new Power Apps app owner using userName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(validUserName)}'&$select=Id`) {
        return userResponse;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${validEnvironmentName}/apps/${validAppName}/modifyAppOwner?api-version=2022-11-01`) {
        return appOwnerSetResponse;
      }
    });

    const requestBody = {
      newAppOwner: 'd2481133-e3ed-4add-836d-6e200969dd03',
      roleForOldAppOwner: undefined
    };

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironmentName, appName: validAppName, userName: validUserName } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('sets a new Power Apps app owner using userId', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${validEnvironmentName}/apps/${validAppName}/modifyAppOwner?api-version=2022-11-01`) {
        return appOwnerSetResponse;
      }
    });

    const requestBody = {
      newAppOwner: 'd2481133-e3ed-4add-836d-6e200969dd03',
      roleForOldAppOwner: undefined
    };

    await command.action(logger, { options: { environmentName: validEnvironmentName, appName: validAppName, userId: validUserId } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('sets a new Power Apps owner using userId and sets old owner with role CanView', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${validEnvironmentName}/apps/${validAppName}/modifyAppOwner?api-version=2022-11-01`) {
        return appOwnerSetResponse;
      }
    });

    const requestBody = {
      newAppOwner: 'd2481133-e3ed-4add-836d-6e200969dd03',
      roleForOldAppOwner: 'CanView'
    };

    await command.action(logger, { options: { environmentName: validEnvironmentName, appName: validAppName, userId: validUserId, roleForOldAppOwner: 'CanView' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('sets a new Power Apps owner using userId and sets old owner with role CanEdit', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${validEnvironmentName}/apps/${validAppName}/modifyAppOwner?api-version=2022-11-01`) {
        return appOwnerSetResponse;
      }
    });

    const requestBody = {
      newAppOwner: 'd2481133-e3ed-4add-836d-6e200969dd03',
      roleForOldAppOwner: 'CanEdit'
    };

    await command.action(logger, { options: { environmentName: validEnvironmentName, appName: validAppName, userId: validUserId, roleForOldAppOwner: 'CanEdit' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = `The specified user with user id ${validUserId} does not exist.`;
    sinon.stub(request, 'post').rejects({
      error: {
        error: {
          message: errorMessage
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: validEnvironmentName, appName: validAppName, userId: validUserId } } as any),
      new CommandError(errorMessage));
  });
});
