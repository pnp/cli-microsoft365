import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import commands from '../../commands.js';
import { sinonUtil } from './../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';
import command from './directoryextension-list.js';

describe(commands.DIRECTORYEXTENSION_LIST, () => {
  const appId = 'fd918e4b-c821-4efb-b50a-5eddd23afc6f';
  const appObjectId = '1caf7dcd-7e83-4c3a-94f7-932a1299c844';
  const appName = 'ContosoApp';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  const response = {
    "value": [
      {
        "id": "8133c498-ad76-4a7b-90a0-675bf5adf492",
        "deletedDateTime": null,
        "appDisplayName": "ContosoApp",
        "dataType": "String",
        "isMultiValued": true,
        "isSyncedFromOnPremises": false,
        "name": "extension_66eac1c505384aec9e024b9e60f5e4b9_jobGroup",
        "targetObjects": [
          "User"
        ]
      }
    ]
  };

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

    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DIRECTORYEXTENSION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'appDisplayName']);
  });

  it('fails validation if appId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if appObjectId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      appObjectId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both appId and appObjectId are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      appObjectId: appObjectId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both appObjectId and appName are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appObjectId: appObjectId,
      appName: appName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both appId and appName are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: appId,
      appName: appName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('retrieves a all available directory extensions', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directoryObjects/getAvailableExtensionProperties`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledWith(response.value));
  });

  it('retrieves a list of directory extensions by appObjectId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appObjectId: appObjectId
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledWith(response.value));
  });

  it('retrieves a list of directory extensions by appId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id`) {
        return {
          "value": [
            {
              "id": appObjectId
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appId: appId
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledWith(response.value));
  });

  it('retrieves a list of directory extensions by appName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '${appName}'&$select=id`) {
        return {
          "value": [
            {
              "id": appObjectId
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appName: appName
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledWith(response.value));
  });

  it('handles error when application specified by appObjectId was not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/`) {
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
      appObjectId: appObjectId
    });

    await assert.rejects(
      command.action(logger, { options: parsedSchema.data }),
      new CommandError(`Resource '${appObjectId}' does not exist or one of its queried reference-property objects are not present.`)
    );
  });
});