import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './enterpriseapp-permission-list.js';

describe(commands.ENTERPRISEAPP_PERMISSION_LIST, () => {
  const appId = 'f1417aa3-bf0b-4cc5-a845-a0b2cf11f690';
  const appName = 'Contoso App';
  const appObjectId = '64381cda-d40d-4fb0-bce2-bece391546a2';
  const graphSpId = '4e8e6f0b-2ba1-4c0e-9c3a-5a1c1b2a7f10';
  const sharePointSpId = '7a2c9a63-6f7b-4b8e-8d0f-3c2e5c1d9e21';
  const defaultAppRoleId = '00000000-0000-0000-0000-000000000000';

  const appRoleAssignmentsUrl = `https://graph.microsoft.com/v1.0/servicePrincipals/${appObjectId}/appRoleAssignments`;
  const oauth2PermissionGrantsUrl = `https://graph.microsoft.com/v1.0/servicePrincipals/${appObjectId}/oauth2PermissionGrants`;
  const graphSpUrl = `https://graph.microsoft.com/v1.0/servicePrincipals/${graphSpId}?$select=id,appId,displayName,appRoles`;
  const sharePointSpUrl = `https://graph.microsoft.com/v1.0/servicePrincipals/${sharePointSpId}?$select=id,appId,displayName,appRoles`;

  //#region Mocked responses
  const graphSp = {
    id: graphSpId,
    appId: '00000003-0000-0000-c000-000000000000',
    displayName: 'Microsoft Graph',
    appRoles: [
      { id: '332a536c-c7ef-4017-ab91-336970924f0d', value: 'Sites.Read.All' },
      { id: 'eb158f57-df43-4751-8b21-b8932adb3d34', value: 'RecordsManagement.ReadWrite.All' }
    ]
  };

  const sharePointSp = {
    id: sharePointSpId,
    appId: '00000003-0000-0ff1-ce00-000000000000',
    displayName: 'Office 365 SharePoint Online',
    appRoles: []
  };

  const appRoleAssignments = [
    {
      id: 'Gm9Lp7ah0EuQ-wh5ci9yURmNLp8uGSlIuXCrpmLDcf4',
      appRoleId: '332a536c-c7ef-4017-ab91-336970924f0d',
      principalId: appObjectId,
      principalType: 'ServicePrincipal',
      resourceDisplayName: 'Microsoft Graph',
      resourceId: graphSpId
    },
    {
      id: 'Gm9Lp7ah0EuQ-wh5ci9yUWJpMV5aHtpOvBpRMoWT7Qc',
      appRoleId: 'eb158f57-df43-4751-8b21-b8932adb3d34',
      principalId: appObjectId,
      principalType: 'ServicePrincipal',
      resourceDisplayName: 'Microsoft Graph',
      resourceId: graphSpId
    }
  ];

  const oauth2PermissionGrants = [
    {
      id: 'aFHtXbGc7UGXzVBbrb6bpdqbHNRTF8xMr-dbQ5Bqgy0',
      clientId: appObjectId,
      consentType: 'AllPrincipals',
      principalId: null,
      resourceId: sharePointSpId,
      scope: ' AllSites.Read '
    },
    {
      id: 'aFHtXbGc7UGXzVBbrb6bpQsOjk6hKw5MnDpaHBsqfxA',
      clientId: appObjectId,
      consentType: 'AllPrincipals',
      principalId: null,
      resourceId: graphSpId,
      scope: 'User.Read openid'
    },
    {
      id: 'aFHtXbGc7UGXzVBbrb6bpQsOjk6hKw5MnDpaHBsqfxB',
      clientId: appObjectId,
      consentType: 'Principal',
      principalId: 'd3a4b2c1-9e8f-4a7b-8c6d-5e4f3a2b1c0d',
      resourceId: graphSpId,
      scope: 'User.Read'
    }
  ];

  const applicationPermissions = [
    {
      resource: 'Microsoft Graph',
      resourceId: '00000003-0000-0000-c000-000000000000',
      permission: 'Sites.Read.All',
      type: 'Application'
    },
    {
      resource: 'Microsoft Graph',
      resourceId: '00000003-0000-0000-c000-000000000000',
      permission: 'RecordsManagement.ReadWrite.All',
      type: 'Application'
    }
  ];

  const delegatedPermissions = [
    {
      resource: 'Office 365 SharePoint Online',
      resourceId: '00000003-0000-0ff1-ce00-000000000000',
      permission: 'AllSites.Read',
      type: 'Delegated'
    },
    {
      resource: 'Microsoft Graph',
      resourceId: '00000003-0000-0000-c000-000000000000',
      permission: 'User.Read',
      type: 'Delegated'
    },
    {
      resource: 'Microsoft Graph',
      resourceId: '00000003-0000-0000-c000-000000000000',
      permission: 'openid',
      type: 'Delegated'
    },
    {
      resource: 'Microsoft Graph',
      resourceId: '00000003-0000-0000-c000-000000000000',
      permission: 'User.Read',
      type: 'Delegated'
    }
  ];
  //#endregion

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const stubGet = (responses: { [url: string]: any }): sinon.SinonStub => {
    return sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url! in responses) {
        return responses[opts.url!];
      }

      throw 'Invalid request';
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENTERPRISEAPP_PERMISSION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ appId: '123' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the appId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ appId: appId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the appObjectId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ appObjectId: '123' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the appObjectId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ appObjectId: appObjectId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if the appName is specified', () => {
    const actual = commandOptionsSchema.safeParse({ appName: appName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if neither the appId, appObjectId, nor appName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ type: 'all' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both the appId and appName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ appId: appId, appName: appName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if appId, appObjectId, and appName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ appId: appId, appObjectId: appObjectId, appName: appName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the type is not a valid permission type', () => {
    const actual = commandOptionsSchema.safeParse({ appId: appId, type: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the type is specified in a different casing', () => {
    const actual = commandOptionsSchema.safeParse({ appId: appId, type: 'Delegated' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ appId: appId, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('lists all application and delegated permissions of the enterprise application specified by appObjectId', async () => {
    stubGet({
      [appRoleAssignmentsUrl]: { value: appRoleAssignments },
      [oauth2PermissionGrantsUrl]: { value: oauth2PermissionGrants },
      [graphSpUrl]: graphSp,
      [sharePointSpUrl]: sharePointSp
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appObjectId: appObjectId }) });
    assert(loggerLogSpy.calledOnceWithExactly([...applicationPermissions, ...delegatedPermissions]));
  });

  it('retrieves each resource service principal only once', async () => {
    const getStub = stubGet({
      [appRoleAssignmentsUrl]: { value: appRoleAssignments },
      [oauth2PermissionGrantsUrl]: { value: oauth2PermissionGrants },
      [graphSpUrl]: graphSp,
      [sharePointSpUrl]: sharePointSp
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appObjectId: appObjectId }) });
    assert.strictEqual(getStub.getCalls().filter(call => call.args[0].url === graphSpUrl).length, 1);
    assert.strictEqual(getStub.getCalls().filter(call => call.args[0].url === sharePointSpUrl).length, 1);
  });

  it('lists all permissions of the enterprise application specified by appId', async () => {
    stubGet({
      [`https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=id`]: { value: [{ id: appObjectId }] },
      [appRoleAssignmentsUrl]: { value: appRoleAssignments },
      [oauth2PermissionGrantsUrl]: { value: oauth2PermissionGrants },
      [graphSpUrl]: graphSp,
      [sharePointSpUrl]: sharePointSp
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appId: appId }) });
    assert(loggerLogSpy.calledOnceWithExactly([...applicationPermissions, ...delegatedPermissions]));
  });

  it('lists all permissions of the enterprise application specified by appName', async () => {
    stubGet({
      [`https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'Contoso%20App'&$select=id`]: { value: [{ id: appObjectId }] },
      [appRoleAssignmentsUrl]: { value: appRoleAssignments },
      [oauth2PermissionGrantsUrl]: { value: oauth2PermissionGrants },
      [graphSpUrl]: graphSp,
      [sharePointSpUrl]: sharePointSp
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appName: appName }) });
    assert(loggerLogSpy.calledOnceWithExactly([...applicationPermissions, ...delegatedPermissions]));
  });

  it('prompts to choose an enterprise application when multiple apps with the specified appName exist', async () => {
    const otherAppObjectId = '9b1b1e0e-2a8c-4b5f-a0f4-1f6a2e3d4c5b';
    stubGet({
      [`https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'Contoso%20App'&$select=id`]: { value: [{ id: otherAppObjectId }, { id: appObjectId }] },
      [appRoleAssignmentsUrl]: { value: appRoleAssignments },
      [oauth2PermissionGrantsUrl]: { value: [] },
      [graphSpUrl]: graphSp
    });
    const promptStub = sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: appObjectId });

    await command.action(logger, { options: commandOptionsSchema.parse({ appName: appName }) });
    assert(promptStub.calledOnce);
    assert(loggerLogSpy.calledOnceWithExactly(applicationPermissions));
  });

  it('lists only application permissions when type is application', async () => {
    const getStub = stubGet({
      [appRoleAssignmentsUrl]: { value: appRoleAssignments },
      [graphSpUrl]: graphSp
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appObjectId: appObjectId, type: 'application' }) });
    assert(loggerLogSpy.calledOnceWithExactly(applicationPermissions));
    assert(getStub.getCalls().every(call => call.args[0].url !== oauth2PermissionGrantsUrl));
  });

  it('lists only delegated permissions when type is delegated', async () => {
    const getStub = stubGet({
      [oauth2PermissionGrantsUrl]: { value: oauth2PermissionGrants },
      [graphSpUrl]: graphSp,
      [sharePointSpUrl]: sharePointSp
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appObjectId: appObjectId, type: 'delegated' }) });
    assert(loggerLogSpy.calledOnceWithExactly(delegatedPermissions));
    assert(getStub.getCalls().every(call => call.args[0].url !== appRoleAssignmentsUrl));
  });

  it('returns the app role id when the app role is not found on the resource', async () => {
    stubGet({
      [appRoleAssignmentsUrl]: {
        value: [
          { ...appRoleAssignments[0], appRoleId: defaultAppRoleId }
        ]
      },
      [graphSpUrl]: graphSp
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appObjectId: appObjectId, type: 'application' }) });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        resource: 'Microsoft Graph',
        resourceId: '00000003-0000-0000-c000-000000000000',
        permission: defaultAppRoleId,
        type: 'Application'
      }
    ]));
  });

  it('returns no delegated permissions for a grant without scopes', async () => {
    stubGet({
      [oauth2PermissionGrantsUrl]: {
        value: [
          { ...oauth2PermissionGrants[1], scope: null },
          { ...oauth2PermissionGrants[2], scope: '' }
        ]
      },
      [graphSpUrl]: graphSp
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appObjectId: appObjectId, type: 'delegated' }) });
    assert(loggerLogSpy.calledOnceWithExactly([]));
  });

  it('returns an empty array when the enterprise application has no permissions', async () => {
    const getStub = stubGet({
      [appRoleAssignmentsUrl]: { value: [] },
      [oauth2PermissionGrantsUrl]: { value: [] }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appObjectId: appObjectId }) });
    assert(loggerLogSpy.calledOnceWithExactly([]));
    assert.strictEqual(getStub.callCount, 2);
  });

  it('retrieves all pages of app role assignments', async () => {
    const nextPageUrl = `${appRoleAssignmentsUrl}?$skiptoken=X%274453707402000100000017`;
    stubGet({
      [appRoleAssignmentsUrl]: { '@odata.nextLink': nextPageUrl, value: [appRoleAssignments[0]] },
      [nextPageUrl]: { value: [appRoleAssignments[1]] },
      [graphSpUrl]: graphSp
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appObjectId: appObjectId, type: 'application' }) });
    assert(loggerLogSpy.calledOnceWithExactly(applicationPermissions));
  });

  it('logs progress in verbose mode when the enterprise application is specified by appId', async () => {
    stubGet({
      [`https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=id`]: { value: [{ id: appObjectId }] },
      [appRoleAssignmentsUrl]: { value: appRoleAssignments },
      [graphSpUrl]: graphSp
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appId: appId, type: 'application', verbose: true }) });
    assert(loggerLogToStderrSpy.calledWith(`Retrieving enterprise application ${appId}...`));
    assert(loggerLogToStderrSpy.calledWith(`Retrieving permissions granted to enterprise application ${appObjectId}...`));
    assert(loggerLogToStderrSpy.calledWith(`Retrieving resource service principal ${graphSpId}...`));
  });

  it('logs progress in verbose mode when the enterprise application is specified by appName', async () => {
    stubGet({
      [`https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'Contoso%20App'&$select=id`]: { value: [{ id: appObjectId }] },
      [oauth2PermissionGrantsUrl]: { value: [] }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ appName: appName, type: 'delegated', verbose: true }) });
    assert(loggerLogToStderrSpy.calledWith(`Retrieving enterprise application ${appName}...`));
  });

  it('throws an error when the enterprise application specified by appId does not exist', async () => {
    stubGet({
      [`https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=id`]: { value: [] }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ appId: appId }) }),
      new CommandError(`Service principal with appId '${appId}' not found in Microsoft Entra ID.`));
  });

  it('throws an error when the enterprise application specified by appName does not exist', async () => {
    stubGet({
      [`https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'Contoso%20App'&$select=id`]: { value: [] }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ appName: appName }) }),
      new CommandError(`Service principal with name '${appName}' not found in Microsoft Entra ID.`));
  });

  it('handles API error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        error: {
          code: 'Request_ResourceNotFound',
          message: `Resource '${appObjectId}' does not exist or one of its queried reference-property objects are not present.`
        }
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ appObjectId: appObjectId }) }),
      new CommandError(`Resource '${appObjectId}' does not exist or one of its queried reference-property objects are not present.`));
  });
});
