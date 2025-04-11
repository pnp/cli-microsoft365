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
import command from './openextension-remove.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import request from '../../../../request.js';
import { CommandError } from '../../../../Command.js';

describe(commands.OPENEXTENSION_REMOVE, () => {
  const resourceId = 'f4099688-dd3f-4a55-a9f5-ddd7417c227a';
  const userPrincipalName = 'john.doe@contoso.com';
  const extensionId = 'com.contoso.roamingSettings';
  let log: any[];
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
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.OPENEXTENSION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if resourceId is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceType: 'user',
      name: 'com.contoso.roamingSettings'
    });

    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if resourceType is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      name: 'com.contoso.roamingSettings'
    });

    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if name is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'user'
    });

    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if resourceId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: 'foo',
      resourceType: 'group',
      name: 'com.contoso.roamingSettings'
    });

    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if resourceType is user and resourceId is neither a valid GUID nor a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: 'foo',
      resourceType: 'user',
      name: 'com.contoso.roamingSettings'
    });

    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if resourceType is not a valid resource type', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'foo',
      name: 'com.contoso.roamingSettings'
    });

    assert.notStrictEqual(actual.success, true);
  });

  it('prompts before removing the open extension when confirm option not passed', async () => {
    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'group',
      name: 'com.contoso.roamingSettings'
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(promptIssued);
  });

  it('aborts removing the open extension when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'group',
      name: 'com.contoso.roamingSettings'
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(deleteSpy.notCalled);
  });

  it('passes validation if resourceType is user and resourceId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'user',
      name: 'com.contoso.roamingSettings'
    });

    assert.strictEqual(actual.success, true);
  });

  it('passes validation if resourceType is user and resourceId is a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: userPrincipalName,
      resourceType: 'user',
      name: 'com.contoso.roamingSettings'
    });

    assert.strictEqual(actual.success, true);
  });

  it('passes validation if resourceType is group', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'group',
      name: 'com.contoso.roamingSettings'
    });

    assert.strictEqual(actual.success, true);
  });

  it('passes validation if resourceType is device', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'device',
      name: 'com.contoso.roamingSettings'
    });

    assert.strictEqual(actual.success, true);
  });

  it('passes validation if resourceType is organization', () => {
    const actual = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'organization',
      name: 'com.contoso.roamingSettings'
    });

    assert.strictEqual(actual.success, true);
  });

  it('removes an open extension defined for a user', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${resourceId}/extensions/${extensionId}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'user',
      name: 'com.contoso.roamingSettings',
      force: true
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes an open extension defined for a group', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${resourceId}/extensions/${extensionId}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'group',
      name: 'com.contoso.roamingSettings',
      verbose: true,
      force: true
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes an open extension defined for a device', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/devices/${resourceId}/extensions/${extensionId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'device',
      name: 'com.contoso.roamingSettings'
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes an open extension defined for an organization', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization/${resourceId}/extensions/${extensionId}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceId: resourceId,
      resourceType: 'organization',
      name: 'com.contoso.roamingSettings',
      force: true
    });

    await command.action(logger, { options: parsedSchema.data });
    assert(deleteRequestStub.calledOnce);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'delete').rejects({
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
      force: true
    });

    await assert.rejects(command.action(logger, { options: parsedSchema.data }), new CommandError('Extension with given id not found.'));
  });
});
