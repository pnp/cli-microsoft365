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
import command from './openextension-add.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import request from '../../../../request.js';
import { CommandError } from '../../../../Command.js';

describe(commands.OPENEXTENSION_ADD, () => {
  const resourceId = 'f4099688-dd3f-4a55-a9f5-ddd7417c227a';
  const response = {
    "@odata.type": "#microsoft.graph.openTypeExtension",
    "extensionName": "com.contoso.roamingSettings",
    "theme": "dark",
    "color": "purple",
    "lang": "Japanese",
    "id": "com.contoso.roamingSettings"
  };
  const responseWithJsonObject = {
    "@odata.type": "#microsoft.graph.openTypeExtension",
    "extensionName": "com.contoso.roamingSettings",
    "supportedSystem": "Linux",
    "settings": {
      "theme": "dark",
      "color": "purple",
      "lang": "Japanese"
    },
    "id": "com.contoso.roamingSettings"
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
    assert.strictEqual(command.name, commands.OPENEXTENSION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if resourceId is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceType: 'user',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if resourceType is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if name is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'user',
      theme: 'dark',
      color: 'red',
      language: 'English'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if resourceId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: 'foo',
      resourceType: 'user',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if resourceType is not a valid resource type', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'foo',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if resourceType is user', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'user',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if resourceType is group', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'group',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if resourceType is device', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'device',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if resourceType is organization', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'organization',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English'
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly creates an open extension defined for a user', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${resourceId}/extensions`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'user',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an open extension defined for a group', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${resourceId}/extensions`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'group',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an open extension defined for a device', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/devices/${resourceId}/extensions`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'device',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an open extension defined for an organization', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization/${resourceId}/extensions`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'organization',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an open extension with JSON object', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${resourceId}/extensions` && 
        JSON.stringify(opts.data) === JSON.stringify({
          'extensionName': 'com.contoso.roamingSettings',
          'settings': {
            "theme": "dark",
            "color": "red",
            "language": "English"
          },
          'supportedSystem': 'Linux'
        })) {
        return responseWithJsonObject;
      }

      throw opts.data;// 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'user',
      name: 'com.contoso.roamingSettings',
      settings: '{"theme": "dark", "color": "red", "language": "English"}',
      supportedSystem: 'Linux',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledOnceWithExactly(responseWithJsonObject));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        error: {
          code: "Request_BadRequest",
          message: "An extension already exists with given id.",
          innerError: {
            date: "2025-03-13T13:15:33",
            'request-id': "f92d9230-3297-4d2b-9ac5-e7b2abc32d4f",
            'client-request-id': "f92d9230-3297-4d2b-9ac5-e7b2abc32d4f"
          }
        }
      }
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'organization',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English',
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data }), new CommandError('An extension already exists with given id.'));
  });
});
