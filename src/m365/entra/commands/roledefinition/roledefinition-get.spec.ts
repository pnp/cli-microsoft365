import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import command from './roledefinition-get.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { z } from 'zod';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.ROLEDEFINITION_GET, () => {
  const roleId = 'abcd1234-de71-4623-b4af-96380a352509';
  const roleDisplayName = 'Bitlocker Keys Reader';
  const roleDefinitionResponse = {
    "id": "abcd1234-de71-4623-b4af-96380a352509",
    "description": "Can read Bitlocker keys.",
    "displayName": "Bitlocker Keys Reader",
    "isBuiltIn": false,
    "isEnabled": true,
    "resourceScopes": [
      "/"
    ],
    "templateId": "abcd1234-de71-4623-b4af-96380a352509",
    "version": "1",
    "rolePermissions": [
      {
        "allowedResourceActions": [
          "microsoft.directory/bitlockerKeys/key/read"
        ],
        "condition": null
      }
    ],
    "inheritsPermissionsFrom": [
    ]
  };

  const roleDefinitionLimitedResponse = {
    "id": "abcd1234-de71-4623-b4af-96380a352509",
    "displayName": "Bitlocker Keys Reader",
    "isBuiltIn": false,
    "isEnabled": true
  };

  let log: string[];
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROLEDEFINITION_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both id and displayName are provided', () => {
    const actual = commandOptionsSchema.safeParse({
      id: roleId,
      displayName: roleDisplayName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither id nor displayName is provided', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it(`should get an Entra ID role definition by id`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${roleId}`) {
        return roleDefinitionResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { id: roleId, verbose: true }
    });

    assert(loggerLogSpy.calledWith(roleDefinitionResponse));
  });

  it(`should get an Entra ID role definition by displayName`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(roleDisplayName)}'`) {
        return { value: [roleDefinitionResponse] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { displayName: roleDisplayName, verbose: true }
    });

    assert(loggerLogSpy.calledWith(roleDefinitionResponse));
  });

  it(`should get an Entra ID role definition by id with specified properties`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${roleId}?$select=id,displayName,isBuiltIn,isEnabled`) {
        return roleDefinitionLimitedResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { id: roleId, properties: 'id,displayName,isBuiltIn,isEnabled' }
    });

    assert(loggerLogSpy.calledWith(roleDefinitionLimitedResponse));
  });

  it(`should get an Entra ID role definition by displayName with specified properties`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(roleDisplayName)}'&$select=id,displayName,isBuiltIn,isEnabled`) {
        return {
          value: [roleDefinitionLimitedResponse]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { displayName: roleDisplayName, properties: 'id,displayName,isBuiltIn,isEnabled' }
    });

    assert(loggerLogSpy.calledWith(roleDefinitionLimitedResponse));
  });

  it('handles error when retrieving role definition failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${roleId}`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: { id: roleId } }),
      new CommandError('An error has occurred')
    );
  });

  it('handles error when role definition was not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${roleId}`) {
        throw {
          error:
          {
            code: 'Request_ResourceNotFound',
            message: `Resource '${roleId}' does not exist or one of its queried reference-property objects are not present.`
          }
        };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: { id: roleId } }),
      new CommandError(`Resource '${roleId}' does not exist or one of its queried reference-property objects are not present.`)
    );
  });
});