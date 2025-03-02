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
import command from './directoryextension-remove.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import request from '../../../../request.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { CommandError } from '../../../../Command.js';
import { directoryExtension } from '../../../../utils/directoryExtension.js';

describe(commands.DIRECTORYEXTENSION_REMOVE, () => {
  const appId = '7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091';
  const appObjectId = '1a70e568-d286-4ad1-b036-734ff8667915';
  const appName = 'ContosoApp';
  const extensionId = '522817ae-5c95-4243-96c1-f85231fcbc1f';
  const extensionName = 'extension_105be60b603845fea385e58772d9d630_githubworkaccount';

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean;
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.promptForConfirmation,
      directoryExtension.getDirectoryExtensionByName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DIRECTORYEXTENSION_REMOVE);
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

  it('prompts before removing the directory extension when confirm option not passed', async () => {
    const parsedSchema = commandOptionsSchema.safeParse({
      appId: appId,
      name: extensionName
    });
    await command.action(logger, { options: parsedSchema.data });

    assert(promptIssued);
  });

  it('aborts removing the directory extension when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    const parsedSchema = commandOptionsSchema.safeParse({
      appId: appId,
      name: extensionName
    });
    await command.action(logger, { options: parsedSchema.data });
    assert(deleteSpy.notCalled);
  });

  it('removes the directory extension specified by id registered for an application specified by appObjectId without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/${extensionId}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appObjectId: appObjectId,
      id: extensionId,
      force: true,
      verbose: true
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(deleteRequestStub.called);
  });

  it('removes the directory extension specified by name registered for an application specified by appId without prompting for confirmation', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves({ id: appObjectId });
    sinon.stub(directoryExtension, 'getDirectoryExtensionByName').resolves({ id: extensionId });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/${extensionId}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      appId: appId,
      name: extensionName,
      force: true,
      verbose: true
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(deleteRequestStub.called);
  });

  it('removes the directory extension specified by name registered for an application specified by name without prompting for confirmation', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppName').resolves({ id: appObjectId });
    sinon.stub(directoryExtension, 'getDirectoryExtensionByName').resolves({ id: extensionId });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties/${extensionId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const parsedSchema = commandOptionsSchema.safeParse({
      appName: appName,
      name: extensionName,
      verbose: true
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(deleteRequestStub.called);
  });

  it('handles error when application specified by id was not found', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
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

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const parsedSchema = commandOptionsSchema.safeParse({
      appObjectId: appObjectId,
      id: extensionId,
      verbose: true
    });

    await assert.rejects(
      command.action(logger, { options: parsedSchema.data }),
      new CommandError(`Resource '${appObjectId}' does not exist or one of its queried reference-property objects are not present.`)
    );
  });

  it('handles error when directory extension specified by id was not found', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
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

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const parsedSchema = commandOptionsSchema.safeParse({
      appObjectId: appObjectId,
      id: extensionId,
      force: true,
      verbose: true
    });

    await assert.rejects(
      command.action(logger, { options: parsedSchema.data }),
      new CommandError(`Resource '${extensionId}' does not exist or one of its queried reference-property objects are not present.`)
    );
  });
});