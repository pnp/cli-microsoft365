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
const command: Command = require('./app-permission-list');

describe(commands.APP_PERMISSION_LIST, () => {
  const environmentName = 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6';
  const appName = '58768250-1943-470b-9743-715020ae21f4';
  const roleName = 'Owner';

  const permissionsResponse = [
    {
      'name': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a',
      'id': '/providers/Microsoft.PowerApps/apps/37ea6004-f07b-46ca-8ef3-a256b67b4dbb/permissions/fe36f75e-c103-410b-a18a-2bf6df06ac3a',
      'type': 'Microsoft.PowerApps/apps/permissions',
      'properties': {
        'roleName': 'Owner',
        'principal': {
          'id': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a',
          'displayName': 'John Doe',
          'email': 'john@contoso.com',
          'type': 'User',
          'tenantId': 'e1dd4023-a656-480a-8a0e-c1b1eec51e1d'
        },
        'scope': '/providers/Microsoft.PowerApps/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/apps/37ea6004-f07b-46ca-8ef3-a256b67b4dbb',
        'notifyShareTargetOption': 'NotSpecified',
        'inviteGuestToTenant': false,
        'createdOn': '2022-10-25T21:28:14.2122305Z',
        'createdBy': 'f0db9c91-3dae-49c8-98fa-8059b8909d45'
      }
    }
  ];

  const permissionsResponseFormatted = [
    {
      'name': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a',
      'id': '/providers/Microsoft.PowerApps/apps/37ea6004-f07b-46ca-8ef3-a256b67b4dbb/permissions/fe36f75e-c103-410b-a18a-2bf6df06ac3a',
      'type': 'Microsoft.PowerApps/apps/permissions',
      'roleName': 'Owner',
      'principalId': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a',
      'principalType': 'User',
      'properties': {
        'roleName': 'Owner',
        'principal': {
          'id': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a',
          'displayName': 'John Doe',
          'email': 'john@contoso.com',
          'type': 'User',
          'tenantId': 'e1dd4023-a656-480a-8a0e-c1b1eec51e1d'
        },
        'scope': '/providers/Microsoft.PowerApps/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/apps/37ea6004-f07b-46ca-8ef3-a256b67b4dbb',
        'notifyShareTargetOption': 'NotSpecified',
        'inviteGuestToTenant': false,
        'createdOn': '2022-10-25T21:28:14.2122305Z',
        'createdBy': 'f0db9c91-3dae-49c8-98fa-8059b8909d45'
      }
    }
  ];


  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_PERMISSION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['roleName', 'principalId', 'principalType']);
  });

  it('correctly retrieves all permissions when asAdmin and environmentName is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}/apps/${appName}/permissions?api-version=2022-11-01`) {
        return { value: permissionsResponse };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: appName, asAdmin: true, environmentName: environmentName, verbose: true } });
    assert(loggerLogSpy.calledWith(permissionsResponseFormatted));
  });

  it('correctly filters permissions when roleName is passed', async () => {
    const permissionsWithDifferentRoleName = { 'name': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a', 'id': '/providers/Microsoft.PowerApps/apps/37ea6004-f07b-46ca-8ef3-a256b67b4dbb/permissions/fe36f75e-c103-410b-a18a-2bf6df06ac3a', 'type': 'Microsoft.PowerApps/apps/permissions', 'properties': { 'roleName': 'CanEdit', 'principal': { 'id': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a', 'displayName': 'John Doe', 'email': 'john@contoso.com', 'type': 'User', 'tenantId': 'e1dd4023-a656-480a-8a0e-c1b1eec51e1d' }, 'scope': '/providers/Microsoft.PowerApps/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/apps/37ea6004-f07b-46ca-8ef3-a256b67b4dbb', 'notifyShareTargetOption': 'NotSpecified', 'inviteGuestToTenant': false, 'createdOn': '2022-10-25T21:28:14.2122305Z', 'createdBy': 'f0db9c91-3dae-49c8-98fa-8059b8909d45' } };
    const permissionsResponseClone = [...permissionsResponse];
    permissionsResponseClone.push(permissionsWithDifferentRoleName);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${appName}/permissions?api-version=2022-11-01`) {
        return { value: permissionsResponseClone };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: appName, roleName: roleName, verbose: true } });
    assert(loggerLogSpy.calledWith(permissionsResponseFormatted));
  });

  it('correctly handles no permissions found', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });

    await command.action(logger, { options: { appName: appName, verbose: true } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('correctly handles no permissions found with specific roleName', async () => {
    const roleName = 'CanEdit';
    sinon.stub(request, 'get').resolves({ value: permissionsResponse });

    await command.action(logger, { options: { appName: appName, roleName: roleName, verbose: true } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('correctly handles API error when app not found or no access', async () => {
    const error = {
      error: {
        code: 'Forbidden',
        message: `The user with object id 'fe36f75e-c103-410b-a18a-2bf6df06ac3a' does not have permission to access this.`
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, { options: { appName: appName } } as any),
      new CommandError(error.error.message));
  });

  it('passes validation if asAdmin specified with environment', async () => {
    const actual = await command.validate({ options: { appName: appName, asAdmin: true, environmentName: environmentName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if roleName is a valid roleName', async () => {
    const actual = await command.validate({ options: { appName: appName, roleName: roleName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if roleName is not a valid roleName', async () => {
    const actual = await command.validate({ options: { appName: appName, roleName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appName is not a valid guid', async () => {
    const actual = await command.validate({ options: { appName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if asAdmin specified without environmentName', async () => {
    const actual = await command.validate({ options: { appName: appName, asAdmin: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if environmentName specified without asAdmin', async () => {
    const actual = await command.validate({ options: { appName: appName, environmentName: environmentName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
