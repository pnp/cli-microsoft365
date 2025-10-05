import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import { options } from './openextension-get.js';
import command from './openextension-set.js';

describe(commands.OPENEXTENSION_SET, () => {
  const resourceId = 'f4099688-dd3f-4a55-a9f5-ddd7417c227a';
  const userPrincipalName = 'john.doe@contoso.com';
  const extensionName = 'com.contoso.roamingSettings';

  let log: any[];
  let logger: Logger;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.OPENEXTENSION_SET);
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
      resourceType: 'group',
      name: 'com.contoso.roamingSettings',
      theme: 'dark',
      color: 'red',
      language: 'English'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if resoruceType is user and resourceId is neiter a valid GUID nor a valid UPN', () => {
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

  it('passes validation if resourceType is user and resourceId is a valid GUID', () => {
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

  it('passes validation if resourceType is user and resourceId is a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: userPrincipalName,
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

  it('correctly updates an open extension defined for a user', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${resourceId}/extensions/${extensionName}`) {
        return {
          extensionName: "com.contoso.roamingSettings",
          theme: 'dark',
          color: 'red',
          language: 'English',
          id: "com.contoso.roamingSettings"
        };
      }

      throw 'Invalid request';
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${resourceId}/extensions/${extensionName}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'user',
      name: 'com.contoso.roamingSettings',
      theme: 'light',
      color: 'blue',
      language: 'Dutch',
      verbose: true
    });

    const requestBody = {
      "@odata.type": "#microsoft.graph.openTypeExtension",
      theme: "light",
      color: "blue",
      language: "Dutch",
      extensionName: "com.contoso.roamingSettings",
      id: "com.contoso.roamingSettings"
    };

    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, requestBody);
  });

  it('does not change values of not specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${resourceId}/extensions/${extensionName}`) {
        return {
          extensionName: "com.contoso.roamingSettings",
          theme: 'dark',
          color: 'red',
          language: 'English',
          id: "com.contoso.roamingSettings"
        };
      }

      throw 'Invalid request';
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${resourceId}/extensions/${extensionName}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'group',
      name: 'com.contoso.roamingSettings',
      color: 'red',
      language: 'Czech',
      keepUnchangedProperties: true,
      verbose: true
    });

    const requestBody = {
      "@odata.type": "#microsoft.graph.openTypeExtension",
      theme: "dark",
      color: "red",
      language: "Czech",
      extensionName: "com.contoso.roamingSettings",
      id: "com.contoso.roamingSettings"
    };
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, requestBody);
  });

  it('correctly clears value of properties of an open extension defined for a device', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/devices/${resourceId}/extensions/${extensionName}`) {
        return {
          extensionName: "com.contoso.roamingSettings",
          theme: 'dark',
          color: 'red',
          language: 'English',
          id: "com.contoso.roamingSettings"
        };
      }

      throw 'Invalid request';
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/devices/${resourceId}/extensions/${extensionName}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'device',
      name: 'com.contoso.roamingSettings',
      theme: '',
      color: '',
      language: 'Dutch',
      verbose: true
    });

    const requestBody = {
      "@odata.type": "#microsoft.graph.openTypeExtension",
      theme: null,
      color: null,
      language: "Dutch",
      extensionName: "com.contoso.roamingSettings",
      id: "com.contoso.roamingSettings"
    };
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, requestBody);
  });

  it('correctly removes the property from an open extension defined for an organization', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization/${resourceId}/extensions/${extensionName}`) {
        return {
          extensionName: "com.contoso.roamingSettings",
          theme: 'dark',
          color: 'red',
          language: 'English',
          id: "com.contoso.roamingSettings"
        };
      }

      throw 'Invalid request';
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization/${resourceId}/extensions/${extensionName}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'organization',
      name: 'com.contoso.roamingSettings',
      color: 'blue',
      language: 'Dutch',
      verbose: true
    });
    const requestBody = {
      "@odata.type": "#microsoft.graph.openTypeExtension",
      color: "blue",
      language: "Dutch",
      extensionName: "com.contoso.roamingSettings",
      id: "com.contoso.roamingSettings"
    };

    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, requestBody);
  });

  it('correctly updates an open extension with JSON object', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${resourceId}/extensions/${extensionName}`) {
        return {
          extensionName: "com.contoso.roamingSettings",
          settings: {
            theme: 'dark',
            color: 'red',
            language: 'English'
          },
          supportedSystem: 'Linux',
          id: "com.contoso.roamingSettings"
        };
      }

      throw 'Invalid request';
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${resourceId}/extensions/${extensionName}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'user',
      name: 'com.contoso.roamingSettings',
      settings: '{"theme": "dark", "color": "blue", "language": "Dutch"}',
      supportedSystem: 'Windows',
      verbose: true
    });

    const requestBody = {
      "@odata.type": "#microsoft.graph.openTypeExtension",
      settings: {
        theme: 'dark',
        color: 'blue',
        language: 'Dutch'
      },
      supportedSystem: 'Windows',
      extensionName: "com.contoso.roamingSettings",
      id: "com.contoso.roamingSettings"
    };

    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, requestBody);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        error: {
          code: "ResourceNotFound",
          message: "Extension with given id not found.",
          innerError: {
            date: "2025-04-07T11:48:13",
            'request-id': "6534c192-7418-421c-bc36-6f38717ae72f",
            'client-request-id': "6534c192-7418-421c-bc36-6f38717ae72f"
          }
        }
      }
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'organization',
      name: 'com.contoso.roamingSettings',
      color: 'blue',
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }), new CommandError('Extension with given id not found.'));
  });
});
