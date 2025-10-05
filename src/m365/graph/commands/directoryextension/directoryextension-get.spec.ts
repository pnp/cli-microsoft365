import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { directoryExtension } from '../../../../utils/directoryExtension.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './directoryextension-get.js';

describe(commands.DIRECTORYEXTENSION_GET, () => {
  const appId = '7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091';
  const appObjectId = '1a70e568-d286-4ad1-b036-734ff8667915';
  const appName = 'ContosoApp';
  const extensionId = '522817ae-5c95-4243-96c1-f85231fcbc1f';
  const extensionName = 'extension_105be60b603845fea385e58772d9d630_githubworkaccount';

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

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      directoryExtension.getDirectoryExtensionByName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DIRECTORYEXTENSION_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if appId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: 'foo',
      name: extensionName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if appObjectId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      appObjectId: 'foo',
      name: extensionName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if appId and appObjectId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      appObjectId: appObjectId,
      name: extensionName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if appId and appName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      appName: appName,
      name: extensionName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if appObjectId and appName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appObjectId: appObjectId,
      appName: appName,
      name: extensionName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      id: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither name nor id is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both name and id are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      id: extensionId,
      name: extensionName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither appId nor appObjectId nor appName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      name: extensionName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('should get the directory extension specified by id registered for an application specified by appObjectId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/${extensionId}`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appObjectId: appObjectId,
      id: extensionId,
      verbose: true
    });

    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledWith(response));
  });

  it('should get the directory extension specified by name registered for an application specified by appId', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves({ id: appObjectId });
    sinon.stub(directoryExtension, 'getDirectoryExtensionByName').resolves({ id: extensionId });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/${extensionId}`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appId: appId,
      name: extensionName,
      verbose: true
    });

    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledWith(response));
  });

  it('should get the directory extension specified by name registered for an application specified by name', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppName').resolves({ id: appObjectId });
    sinon.stub(directoryExtension, 'getDirectoryExtensionByName').resolves({ id: extensionId });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/${extensionId}`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appName: appName,
      name: extensionName,
      verbose: true
    });

    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledWith(response));
  });

  it('handles error when application specified by id was not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/${extensionId}`) {
        throw {
          error:
          {
            code: 'Request_ResourceNotFound',
            message: `Resource '${appObjectId}' does not exist or one of its queried reference-property objects are not present.`
          }
        };
      }
      throw `Invalid request`;
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appObjectId: appObjectId,
      id: extensionId,
      verbose: true
    });

    await assert.rejects(
      command.action(logger, { options: parsedSchema.data! }),
      new CommandError(`Resource '${appObjectId}' does not exist or one of its queried reference-property objects are not present.`)
    );
  });

  it('handles error when directory extension specified by id was not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/${extensionId}`) {
        throw {
          error:
          {
            code: 'Request_ResourceNotFound',
            message: `Resource '${extensionId}' does not exist or one of its queried reference-property objects are not present.`
          }
        };
      }
      throw `Invalid request`;
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appObjectId: appObjectId,
      id: extensionId,
      verbose: true
    });

    await assert.rejects(
      command.action(logger, { options: parsedSchema.data! }),
      new CommandError(`Resource '${extensionId}' does not exist or one of its queried reference-property objects are not present.`)
    );
  });
});