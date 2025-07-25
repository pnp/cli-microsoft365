import assert from 'assert';
import Configstore from 'configstore';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import cliConfig from '../../../../config.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { cli } from '../../../../cli/cli.js';
import command from './app-add.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { accessToken } from '../../../../utils/accessToken.js';
import request from '../../../../request.js';

describe(commands.APP_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let config: Configstore;
  let configSetSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    config = cli.getConfig();
    configSetSpy = sinon.stub(config, 'set').returns();
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
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.getTenantIdFromAccessToken,
      entraApp.resolveApis,
      entraApp.createAppRegistration,
      entraApp.grantAdminConsent
    ]);
    configSetSpy.resetHistory();
  });

  after(() => {
    sinon.restore();
    sinonUtil.restore([
      config.set
    ]);
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if scopes is not a valid scope option', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Custom App',
      scopes: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if scopes is minimal', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Custom App',
      scopes: 'minimal'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if scopes is all', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Custom App',
      scopes: 'all'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if scopes contains list of scopes', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Custom App',
      scopes: 'https://graph.microsoft.com/User.Read,https://graph.microsoft.com/Group.Read'
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly creates an app registration with all scopes and custom name without saving the app registration info to the CLI config', async () => {
    sinon.stub(accessToken, 'getTenantIdFromAccessToken').returns('00000000-0000-0000-0000-000000000003');
    const scopes = [
      {
        resourceAppId: '00000000-0000-0000-0000-000000000000',
        resourceAccess: [
          {
            id: '00000000-0000-0000-0000-000000000000',
            type: 'Minimal'
          }
        ]
      }
    ];
    sinon.stub(entraApp, 'resolveApis').resolves(scopes);
    const createAppRegistrationSpy = sinon.stub(entraApp, 'createAppRegistration').resolves({
      appId: '00000000-0000-0000-0000-000000000001',
      id: '00000000-0000-0000-0000-000000000002',
      tenantId: '00000000-0000-0000-0000-000000000003',
      requiredResourceAccess: scopes
    });
    sinon.stub(entraApp, 'grantAdminConsent').resolves();
    const parsedSchema = commandOptionsSchema.safeParse({
      name: 'Custom App',
      scopes: 'all',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data });
    assert.deepEqual(createAppRegistrationSpy.getCall(0).args[0], {
      options: {
        allowPublicClientFlows: true,
        apisDelegated: cliConfig.allScopes.join(','),
        implicitFlow: false,
        multitenant: false,
        name: 'Custom App',
        platform: 'publicClient',
        redirectUris: 'http://localhost,https://localhost,https://login.microsoftonline.com/common/oauth2/nativeclient'
      },
      unknownOptions: {},
      apis: scopes,
      logger: logger,
      verbose: true,
      debug: false
    });
  });

  it('correctly creates an app registration with default scopes and default name and saves the app registration info to the CLI config', async () => {
    sinon.stub(accessToken, 'getTenantIdFromAccessToken').returns('00000000-0000-0000-0000-000000000003');
    const scopes = [
      {
        resourceAppId: '00000000-0000-0000-0000-000000000000',
        resourceAccess: [
          {
            id: '00000000-0000-0000-0000-000000000000',
            type: 'Minimal'
          }
        ]
      }
    ];
    sinon.stub(entraApp, 'resolveApis').resolves(scopes);
    const createAppRegistrationSpy = sinon.stub(entraApp, 'createAppRegistration').resolves({
      appId: '00000000-0000-0000-0000-000000000001',
      id: '00000000-0000-0000-0000-000000000002',
      tenantId: '00000000-0000-0000-0000-000000000003',
      requiredResourceAccess: scopes
    });
    sinon.stub(entraApp, 'grantAdminConsent').resolves();
    const parsedSchema = commandOptionsSchema.safeParse({
      saveToConfig: true,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data });
    const expected = {
      clientId: '00000000-0000-0000-0000-000000000001',
      tenantId: '00000000-0000-0000-0000-000000000003'
    };
    assert.deepEqual(createAppRegistrationSpy.getCall(0).args[0], {
      options: {
        allowPublicClientFlows: true,
        apisDelegated: cliConfig.minimalScopes.join(','),
        implicitFlow: false,
        multitenant: false,
        name: 'CLI for M365',
        platform: 'publicClient',
        redirectUris: 'http://localhost,https://localhost,https://login.microsoftonline.com/common/oauth2/nativeclient'
      },
      unknownOptions: {},
      apis: scopes,
      logger: logger,
      verbose: true,
      debug: false
    });
    Object.keys(expected).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (expected as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('correctly creates an app registration with list of scopes and custom name without saving the app registration info to the CLI config', async () => {
    sinon.stub(accessToken, 'getTenantIdFromAccessToken').returns('00000000-0000-0000-0000-000000000003');
    const scopes = [
      {
        resourceAppId: '00000003-0000-0000-c000-000000000000',
        resourceAccess: [
          {
            id: 'e1fe6dd8-ba31-4d61-89e7-88639da4683d',
            type: 'Scope'
          },
          {
            id: '5f8c59db-677d-491f-a6b8-5f174b11ec1d',
            type: 'Scope'
          }
        ]
      }
    ];
    sinon.stub(entraApp, 'resolveApis').resolves(scopes);
    const createAppRegistrationSpy = sinon.stub(entraApp, 'createAppRegistration').resolves({
      appId: '00000000-0000-0000-0000-000000000001',
      id: '00000000-0000-0000-0000-000000000002',
      tenantId: '00000000-0000-0000-0000-000000000003',
      requiredResourceAccess: scopes
    });
    sinon.stub(entraApp, 'grantAdminConsent').resolves();
    const parsedSchema = commandOptionsSchema.safeParse({
      name: 'Custom App',
      scopes: 'https://graph.microsoft.com/User.Read,https://graph.microsoft.com/Group.Read.All',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data });
    assert.deepEqual(createAppRegistrationSpy.getCall(0).args[0], {
      options: {
        allowPublicClientFlows: true,
        apisDelegated: 'https://graph.microsoft.com/User.Read,https://graph.microsoft.com/Group.Read.All',
        implicitFlow: false,
        multitenant: false,
        name: 'Custom App',
        platform: 'publicClient',
        redirectUris: 'http://localhost,https://localhost,https://login.microsoftonline.com/common/oauth2/nativeclient'
      },
      unknownOptions: {},
      apis: scopes,
      logger: logger,
      verbose: true,
      debug: false
    });
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(accessToken, 'getTenantIdFromAccessToken').returns('00000000-0000-0000-0000-000000000003');
    const scopes = [
      {
        resourceAppId: '00000000-0000-0000-0000-000000000000',
        resourceAccess: [
          {
            id: '00000000-0000-0000-0000-000000000000',
            type: 'Minimal'
          }
        ]
      }
    ];
    sinon.stub(entraApp, 'resolveApis').resolves(scopes);
    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });
    const parsedSchema = commandOptionsSchema.safeParse({
      saveToConfig: true,
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data }), new CommandError('Invalid request'));
  });
});
