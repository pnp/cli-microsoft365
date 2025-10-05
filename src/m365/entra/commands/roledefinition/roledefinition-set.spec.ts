import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './roledefinition-set.js';

describe(commands.ROLEDEFINITION_SET, () => {
  const roleId = 'abcd1234-de71-4623-b4af-96380a352509';
  const roleDisplayName = 'Custom Role';

  let log: string[];
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
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROLEDEFINITION_SET);
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

  it('fails validation if neither newDisplayName, description, allowedResourceActions, enabled nor version is provided', () => {
    const actual = commandOptionsSchema.safeParse({ id: roleId });
    assert.notStrictEqual(actual.success, true);
  });

  it('updates a custom role definition specified by id', async () => {
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${roleId}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({ id: roleId, allowedResourceActions: "microsoft.directory/groups.unified/create,microsoft.directory/groups.unified/delete" });
    await command.action(logger, { options: parsedSchema.data! });
    assert(patchRequestStub.called);
  });

  it('updates a custom role definition specified by displayName', async () => {
    sinon.stub(roleDefinition, 'getRoleDefinitionByDisplayName').resolves({ id: roleId });

    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${roleId}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      displayName: roleDisplayName,
      newDisplayName: 'Custom Role Test',
      description: 'Allows creating and deleting unified groups',
      allowedResourceActions: "microsoft.directory/groups.unified/create,microsoft.directory/groups.unified/delete",
      enabled: false,
      version: "2",
      verbose: true
    });
    await command.action(logger, {
      options: parsedSchema.data!
    });
    assert(patchRequestStub.called);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'patch').rejects({
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
      displayName: 'Custom Role',
      allowedResourceActions: "microsoft.directory/groups.unified/create"
    });
    await assert.rejects(command.action(logger, {
      options: parsedSchema.data!
    }), new CommandError('Invalid request'));
  });
});