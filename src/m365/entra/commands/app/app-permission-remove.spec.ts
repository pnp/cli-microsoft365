import { Application, ServicePrincipal } from '@microsoft/microsoft-graph-types';
import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { odata } from '../../../../utils/odata.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-permission-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APP_PERMISSION_REMOVE, () => {
  const appId = '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d';
  const appObjectId = '2aaf2d9e-815e-4a3e-bb80-9b3d9c79078b';
  const appName = 'Dummy Application';

  const servicePrincipalId = '7c330108-8825-4b6c-b280-8d1d68da6bd7';
  const servicePrincipals: ServicePrincipal[] = [{ appId: appId, id: servicePrincipalId, 'servicePrincipalNames': [] }, { "appId": "00000003-0000-0000-c000-000000000000", "id": "fb4be1df-eaa6-4bd0-a068-71f9b2cbe2be", "servicePrincipalNames": ["https://canary.graph.microsoft.com/", "https://graph.microsoft.us/", "https://dod-graph.microsoft.us/", "00000003-0000-0000-c000-000000000000/ags.windows.net", "00000003-0000-0000-c000-000000000000", "https://canary.graph.microsoft.com", "https://graph.microsoft.com", "https://ags.windows.net", "https://graph.microsoft.us", "https://graph.microsoft.com/", "https://dod-graph.microsoft.us"], "appRoles": [{ "allowedMemberTypes": ["Application"], "description": "Allows the app to read and update user profiles without a signed in user.", "displayName": "Read and write all users' full profiles", "id": "741f803b-c850-494e-b5df-cde7c675a1ca", "isEnabled": true, "origin": "Application", "value": "User.ReadWrite.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read user profiles without a signed in user.", "displayName": "Read all users' full profiles", "id": "df021288-bdef-4463-88db-98f22de89214", "isEnabled": true, "origin": "Application", "value": "User.Read.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read and query your audit log activities, without a signed-in user.", "displayName": "Read all audit log data", "id": "b0afded3-3588-46d8-8b3d-9842eff778da", "isEnabled": true, "origin": "Application", "value": "AuditLog.Read.All" }], "oauth2PermissionScopes": [{ "adminConsentDescription": "Allows the app to see and update the data you gave it access to, even when users are not currently using the app. This does not give the app any additional permissions.", "adminConsentDisplayName": "Maintain access to data you have given it access to", "id": "7427e0e9-2fba-42fe-b0c0-848c9e6a8182", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to see and update the data you gave it access to, even when you are not currently using the app. This does not give the app any additional permissions.", "userConsentDisplayName": "Maintain access to data you have given it access to", "value": "offline_access" }, { "adminConsentDescription": "Allows the app to read the available Teams templates, on behalf of the signed-in user.", "adminConsentDisplayName": "Read available Teams templates", "id": "cd87405c-5792-4f15-92f7-debc0db6d1d6", "isEnabled": true, "type": "User", "userConsentDescription": "Read available Teams templates, on your behalf.", "userConsentDisplayName": "Read available Teams templates", "value": "TeamTemplates.Read" }] }];
  const applications: Application[] = [{ id: appObjectId, appId: appId, requiredResourceAccess: [{ resourceAppId: "00000003-0000-0000-c000-000000000000", resourceAccess: [{ id: "e4aa47b9-9a69-4109-82ed-36ec70d85ff1", type: "Scope" }, { id: "7427e0e9-2fba-42fe-b0c0-848c9e6a8182", type: "Scope" }, { id: "332a536c-c7ef-4017-ab91-336970924f0d", type: "Role" }] }] }];
  const applicationPermissions = 'https://graph.microsoft.com/User.ReadWrite.All https://graph.microsoft.com/User.Read.All';
  const delegatedPermissions = 'https://graph.microsoft.com/offline_access';
  const selectProperties = '$select=id,appId,requiredResourceAccess';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      request.post,
      request.delete,
      odata.getAllItems,
      cli.getSettingWithDefaultValue,
      cli.promptForConfirmation,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_PERMISSION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the app when force option not passed', async () => {
    await command.action(logger, { options: { appId: appId, applicationPermissions: applicationPermissions } });
    assert(promptIssued);
  });

  it('aborts removing the app when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { appId: appId, applicationPermissions: applicationPermissions, verbose: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('deletes application permissions and prompt for specifying application when multiple applications found by name and revokes admin consent', async () => {
    const appRoleAssignmentsResponse = [
      {
        id: 'P-xE0cQGikiP9FoIACRlwUa883F-Po5OvrTyGaYOliU',
        appRoleId: '741f803b-c850-494e-b5df-cde7c675a1ca',
        resourceId: 'fb4be1df-eaa6-4bd0-a068-71f9b2cbe2be'
      },
      {
        id: 'P-xE0cQGikiP9FoIACRlwfMbXDqFcFhOncNK1vAO2fU',
        appRoleId: '332a536c-c7ef-4017-ab91-336970924f0d',
        resourceId: 'fb4be1df-eaa6-4bd0-a068-71f9b2cbe2be'
      },
      {
        id: 'P-xE0cQGikiP9FoIACRlwe75C0BIIThEmBbIeQzeWU8',
        appRoleId: 'df021288-bdef-4463-88db-98f22de89214',
        resourceId: 'fb4be1df-eaa6-4bd0-a068-71f9b2cbe2be'
      }
    ];

    const applicationsCopy = [...applications];
    applicationsCopy.push({ id: '340a4aa3-1af6-43ac-87d8-189819003952' });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    sinon.stub(cli, 'handleMultipleResultsFound').resolves(applicationsCopy[0]);

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      switch (url) {
        case `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '${appName}'&${selectProperties}`:
          return applicationsCopy;
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments?$select=id,appRoleId,resourceId`:
          return appRoleAssignmentsResponse;
        default:
          throw 'Invalid request';
      }
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "e4aa47b9-9a69-4109-82ed-36ec70d85ff1",
                  "type": "Scope"
                },
                {
                  "id": "7427e0e9-2fba-42fe-b0c0-848c9e6a8182",
                  "type": "Scope"
                },
                {
                  "id": "332a536c-c7ef-4017-ab91-336970924f0d",
                  "type": "Role"
                }
              ]
            }
          ]
        })) {
        return;
      }
      throw 'Invalid request';
    });

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments/${appRoleAssignmentsResponse[0].id}`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments/${appRoleAssignmentsResponse[2].id}`:
          return;
        default:
          throw 'Invalid request';
      }
    });

    await command.action(logger, { options: { appName: appName, applicationPermissions: applicationPermissions, revokeAdminConsent: true, verbose: true } });
    assert(deleteStub.calledTwice);
  });

  it('deletes delegated permissions from app specified by appObjectId and revokes admin consent', async () => {
    const oAuth2GrantResponse = [
      {
        id: 'P-xE0cQGikiP9FoIACRlwd_hS_um6tBLoGhx-bLL4r4',
        resourceId: 'fb4be1df-eaa6-4bd0-a068-71f9b2cbe2be',
        scope: 'offline_access AgreementAcceptance.Read'
      }
    ];

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}?${selectProperties}`) {
        return applications[0];
      }

      throw 'Invalid request';
    });

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipalId}/oAuth2PermissionGrants?$select=id,resourceId,scope`:
          return oAuth2GrantResponse;
        default:
          throw 'Invalid request';
      }
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/applications/${appObjectId}`:
          if (JSON.stringify(opts.data) === JSON.stringify({
            "requiredResourceAccess": [
              {
                "resourceAppId": "00000003-0000-0000-c000-000000000000",
                "resourceAccess": [
                  {
                    "id": "e4aa47b9-9a69-4109-82ed-36ec70d85ff1",
                    "type": "Scope"
                  },
                  {
                    "id": "332a536c-c7ef-4017-ab91-336970924f0d",
                    "type": "Role"
                  }
                ]
              }
            ]
          })) {
            return;
          }
          else {
            throw 'Invalid request';
          }
        case `https://graph.microsoft.com/v1.0/oauth2PermissionGrants/${oAuth2GrantResponse[0].id}`:
          return;
        default:
          throw 'Invalid request';
      }
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { appObjectId: appObjectId, delegatedPermissions: delegatedPermissions, revokeAdminConsent: true, verbose: true } });
    assert(patchStub.lastCall.args[0].data.scope === 'AgreementAcceptance.Read');
  });

  it('deletes delegated permissions from app specified by appId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}?${selectProperties}`) {
        return applications[0];
      }

      throw 'Invalid request';
    });

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      switch (url) {
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&${selectProperties}`:
          return applications;
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments?$select=id,appRoleId,resourceId`:
          return [];
        default:
          throw 'Invalid request';
      }
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "e4aa47b9-9a69-4109-82ed-36ec70d85ff1",
                  "type": "Scope"
                },
                {
                  "id": "332a536c-c7ef-4017-ab91-336970924f0d",
                  "type": "Role"
                }
              ]
            }
          ]
        })) {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: appObjectId });

    await command.action(logger, { options: { appId: appId, delegatedPermissions: delegatedPermissions, force: true, verbose: true } });
    assert(patchStub.calledOnce);
  });

  it('submits empty array when removing last delegated permission', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}?${selectProperties}`) {
        return applications[0];
      }

      throw 'Invalid request';
    });

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      switch (url) {
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&${selectProperties}`:
          return [{ id: appObjectId, appId: appId, requiredResourceAccess: [{ resourceAppId: "00000003-0000-0000-c000-000000000000", resourceAccess: [{ id: "cd87405c-5792-4f15-92f7-debc0db6d1d6", type: "Scope" }, { id: "7427e0e9-2fba-42fe-b0c0-848c9e6a8182", type: "Scope" }] }] }];
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments?$select=id,appRoleId,resourceId`:
          return [];
        default:
          throw 'Invalid request';
      }
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}`) {
        if (opts.data.requiredResourceAccess.length === 0) {
          return;
        }
        else {
          throw 'Invalid request. Expected empty requiredResourceAccess array';
        }
      }
      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: appObjectId });

    await command.action(logger, { options: { appId: appId, delegatedPermissions: 'https://graph.microsoft.com/TeamTemplates.Read https://graph.microsoft.com/offline_access', force: true, verbose: true } });
    assert(patchStub.calledOnce);
  });

  it('submits empty array when removing last application permission', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}?${selectProperties}`) {
        return applications[0];
      }

      throw 'Invalid request';
    });

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      switch (url) {
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&${selectProperties}`:
          return [{ id: appObjectId, appId: appId, requiredResourceAccess: [{ resourceAppId: "00000003-0000-0000-c000-000000000000", resourceAccess: [{ id: "741f803b-c850-494e-b5df-cde7c675a1ca", type: "Role" }] }] }];
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments?$select=id,appRoleId,resourceId`:
          return [];
        default:
          throw 'Invalid request';
      }
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}`) {
        if (opts.data.requiredResourceAccess.length === 0) {
          return;
        }
        else {
          throw 'Invalid request. Expected empty requiredResourceAccess array';
        }
      }
      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: appObjectId });

    await command.action(logger, { options: { appId: appId, applicationPermissions: 'https://graph.microsoft.com/User.ReadWrite.All', force: true, verbose: true } });
    assert(patchStub.calledOnce);
  });

  it('throws error if application specified by name cannot be found', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '${appName}'&${selectProperties}`) {
        return [];
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { verbose: true, appName: appName, delegatedPermissions: delegatedPermissions, force: true } }),
      new CommandError(`App with name ${appName} not found in Microsoft Entra ID`));
  });

  it('throws error if application specified by appId cannot be found', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&${selectProperties}`) {
        return [];
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { verbose: true, appId: appId, delegatedPermissions: delegatedPermissions, force: true } }),
      new CommandError(`App with id ${appId} not found in Microsoft Entra ID`));
  });

  it('throws an error when service principal is not found', async () => {
    const applicationPermission = 'https://grax.microsoft.com/User.ReadWrite.All';
    const pos: number = applicationPermission.lastIndexOf('/');
    const servicePrincipalName: string = applicationPermission.substring(0, pos);
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id,appId,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, { options: { appId: appId, applicationPermissions: applicationPermission, verbose: true, force: true } }),
      new CommandError(`Service principal ${servicePrincipalName} not found`));
  });

  it('throws an error when permission is not found', async () => {
    const applicationPermission = 'https://graph.microsoft.com/NotFound.All';
    const pos: number = applicationPermission.lastIndexOf('/');
    const servicePrincipalName: string = applicationPermission.substring(0, pos);
    const permissionName: string = applicationPermission.substring(pos + 1);
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id,appId,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, { options: { appId: appId, applicationPermissions: applicationPermission, verbose: true, force: true } }),
      new CommandError(`Permission ${permissionName} for service principal ${servicePrincipalName} not found`));
  });

  it('fails validation if the appId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: 'invalid', applicationPermissions: applicationPermissions } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the appId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: appId, applicationPermissions: applicationPermissions } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the appObjectId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appObjectId: 'invalid', applicationPermissions: applicationPermissions } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the appObjectId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appObjectId: appObjectId, applicationPermissions: applicationPermissions } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});