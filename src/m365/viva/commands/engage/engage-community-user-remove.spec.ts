
import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './engage-community-user-remove.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { entraUser } from '../../../../utils/entraUser.js';

describe(commands.ENGAGE_COMMUNITY_USER_REMOVE, () => {
  const communityId = 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIzNjAyMDAxMTAwOSJ9';
  const communityDisplayName = 'All company';
  const entraGroupId = 'b6c35b51-ebca-445c-885a-63a67d24cb53';
  const userName = 'john@contoso.com';
  const userId = '3f2504e0-4f89-11d3-9a0c-0305e82c3301';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    sinon.stub(entraUser, 'getUserIdByUpn').resolves(userId);
    sinon.stub(vivaEngage, 'getEntraGroupIdByCommunityDisplayName').resolves(entraGroupId);
    sinon.stub(vivaEngage, 'getEntraGroupIdByCommunityId').resolves(entraGroupId);
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
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_COMMUNITY_USER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if entraGroupId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      entraGroupId: 'invalid',
      userName: userName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      entraGroupId: entraGroupId,
      id: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is invalid user principal name', () => {
    const actual = commandOptionsSchema.safeParse({
      entraGroupId: entraGroupId,
      userName: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if communityId, communityDisplayName or entraGroupId are not specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if communityId, communityDisplayName and entraGroupId are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityId: communityId,
      communityDisplayName: communityDisplayName,
      entraGroupId: entraGroupId,
      id: userId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if communityId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityId: communityId,
      userName: userName
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if entraGroupId is specified with a proper GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      entraGroupId: entraGroupId,
      userName: userName
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if communityDisplayName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityDisplayName: communityDisplayName,
      userName: userName
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly removes user specified by id', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/owners/${userId}/$ref`) {
        return;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/members/${userId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { communityDisplayName: communityDisplayName, id: userId, force: true, verbose: true } });
    assert(deleteStub.calledTwice);
  });

  it('correctly removes user by userName', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/owners/${userId}/$ref`) {
        return;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/members/${userId}/$ref`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { communityId: communityId, verbose: true, userName: userName, force: true } });
    assert(deleteStub.calledTwice);
  });

  it('correctly removes user as member by userName', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/owners/${userId}/$ref`) {
        throw {
          response: {
            status: 404,
            data: {
              message: 'Object does not exist...'
            }
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/members/${userId}/$ref`) {
        return;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { communityId: communityId, verbose: true, userName: userName } });
    assert(deleteStub.calledTwice);
  });

  it('handles API error when removing user', async () => {
    const errorMessage = 'Invalid object identifier';
    sinon.stub(request, 'delete').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${entraGroupId}/owners/${userId}/$ref`) {
        throw {
          response: {
            status: 400,
            data: { error: { 'odata.error': { message: { value: errorMessage } } } }
          }
        };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, { options: { entraGroupId: entraGroupId, id: userId } }),
      new CommandError(errorMessage));
  });

  it('prompts before removal when confirmation argument not passed', async () => {
    const promptStub: sinon.SinonStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { entraGroupId: entraGroupId, id: userId } });

    assert(promptStub.called);
  });

  it('aborts execution when prompt not confirmed', async () => {
    const deleteStub = sinon.stub(request, 'delete');
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { entraGroupId: entraGroupId, id: userId } });
    assert(deleteStub.notCalled);
  });
});