import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { cli } from '../../../../cli/cli.js';
import command from './directoryextension-add.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import request from '../../../../request.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { CommandError } from '../../../../Command.js';

describe(commands.DIRECTORYEXTENSION_ADD, () => {
  const appId = '7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091';
  const appObjectId = '1a70e568-d286-4ad1-b036-734ff8667915';
  const appName = 'ContosoApp';
  const response = {
    "id": "522817ae-5c95-4243-96c1-f85231fcbc1f",
    "deletedDateTime": null,
    "appDisplayName": "ContosoApp",
    "dataType": "String",
    "isMultiValued": false,
    "isSyncedFromOnPremises": false,
    "name": "extension_105be60b603845fea385e58772d9d630_GitHubWorkAccount",
    "targetObjects": [
      "User"
    ]
  };
  const responseForMultiValued = {
    "id": "522817ae-5c95-4243-96c1-f85231fcbc1f",
    "deletedDateTime": null,
    "appDisplayName": "ContosoApp",
    "dataType": "String",
    "isMultiValued": true,
    "isSyncedFromOnPremises": false,
    "name": "extension_105be60b603845fea385e58772d9d630_GitHubAccounts",
    "targetObjects": [
      "User"
    ]
  };
  const responseWithMultipleTargets = {
    "id": "522817ae-5c95-4243-96c1-f85231fcbc1f",
    "deletedDateTime": null,
    "appDisplayName": "ContosoApp",
    "dataType": "Boolean",
    "isMultiValued": false,
    "isSyncedFromOnPremises": false,
    "name": "extension_105be60b603845fea385e58772d9d630_ForServiceUseOnly",
    "targetObjects": [
      "User",
      "Application",
      "Device"
    ]
  };
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DIRECTORYEXTENSION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if appId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: 'foo',
      name: 'GitHubWorkAccount',
      dataType: 'String',
      targetObjects: 'User'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if appObjectId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      appObjectId: 'foo',
      name: 'GitHubWorkAccount',
      dataType: 'String',
      targetObjects: 'User'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if dataType is not a valid enum value', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      name: 'GitHubWorkAccount',
      dataType: 'foo',
      targetObjects: 'User'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if targetObjects is not a valid enum value', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      name: 'GitHubWorkAccount',
      dataType: 'String',
      targetObjects: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if targetObjects has more values and one is not a valid enum value', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      name: 'GitHubWorkAccount',
      dataType: 'String',
      targetObjects: 'User,foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if appId and appObjectId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      appObjectId: appObjectId,
      name: 'GitHubWorkAccount',
      dataType: 'String',
      targetObjects: 'User'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if appId and appName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      appName: appName,
      name: 'GitHubWorkAccount',
      dataType: 'String',
      targetObjects: 'User'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if appObjectId and appName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appObjectId: appObjectId,
      appName: appName,
      name: 'GitHubWorkAccount',
      dataType: 'String',
      targetObjects: 'User'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if name is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      dataType: 'String',
      targetObjects: 'User'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if dataType is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      name: 'GitHubWorkAccount',
      targetObjects: 'User'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if targetObjects is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      name: 'GitHubWorkAccount',
      dataType: 'String'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither appId nor appObjectId nor appName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'GitHubWorkAccount',
      dataType: 'String',
      targetObjects: 'User'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('correctly creates a directory extension defined on the application specified by appObjectId', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appObjectId: appObjectId,
      name: 'GitHubWorkAccount',
      dataType: 'String',
      targetObjects: 'User',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates a directory extension defined on the application specified by appId', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves({ id: appObjectId });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties`) {
        return responseForMultiValued;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appId: appId,
      name: 'GitHubAccounts',
      dataType: 'String',
      targetObjects: 'User',
      isMultiValued: true
    });
    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledOnceWithExactly(responseForMultiValued));
  });

  it('correctly creates a directory extension defined on the application specified by appName', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppName').resolves({ id: appObjectId });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties`) {
        return responseWithMultipleTargets;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appName: appName,
      name: 'ForServiceUseOnly',
      dataType: 'Boolean',
      targetObjects: 'User,Application,Device'
    });
    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledOnceWithExactly(responseWithMultipleTargets));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          code: 'Request_MultipleObjectsWithSameKeyValue',
          message: {
            value: 'An extension property exists with the name extension_7f5df2f49ed64df786d7eefbfc4ab091_ForServiceUseOnly.'
          }
        }
      }
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appId: appId,
      name: 'ForServiceUseOnly',
      dataType: 'Boolean',
      targetObjects: 'User,Application,Device'
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data }), new CommandError('An extension property exists with the name extension_7f5df2f49ed64df786d7eefbfc4ab091_ForServiceUseOnly.'));
  });
});