import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
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
import command from './m365group-user-remove.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';

describe(commands.M365GROUP_USER_REMOVE, () => {
  const userName = 'adelev@contoso.com';
  const groupOrTeamId = '80ecb711-2501-4262-b29a-838d30bd3387';
  const userId = '8b38aeff-1642-47e4-b6ef-9d50d29638b7';
  const groupOrTeamName = 'Project Team';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves('00000000-0000-0000-0000-000000000000');
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });
    promptIssued = false;
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(true);
    sinon.stub(entraUser, 'getUserIdsByUpns').resolves([userId]);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue,
      entraUser.getUserIdsByUpns,
      entraGroup.isUnifiedGroup,
      entraGroup.getGroupIdByDisplayName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_USER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the groupId is not a valid guid', async () => {
    const actual = commandOptionsSchema.safeParse({
      groupId: 'invalid',
      userNames: userName
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: 'invalid',
      userNames: userName
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if ids contain an invalid guid', async () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: groupOrTeamId,
      ids: `invalid,${userId}`
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if userNames contain an invalid upn', async () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: groupOrTeamId,
      userNames: `invalid,${userName}`
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when a valid teamId and userNames are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: groupOrTeamId,
      userNames: `${userName},john@contoso.com`
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when a valid teamId and ids are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: groupOrTeamId,
      ids: `${userId},8b38aeff-1642-47e4-b6ef-9d50d29638b7`
    });
    assert.strictEqual(actual.success, true);
  });


  it('prompts before removing the specified user from the specified Microsoft 365 Group when force option not passed', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: userName }) });

    assert(promptIssued);
  });

  it('prompts before removing the specified user from the specified Team when force option not passed (debug)', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, teamId: "00000000-0000-0000-0000-000000000000", userNames: userName }) });

    assert(promptIssued);
  });

  it('aborts removing the specified user from the specified Microsoft 365 Group when force option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: userName }) });
    assert(postSpy.notCalled);
  });

  it('aborts removing the specified user from the specified Microsoft 365 Group when force option not passed and prompt not confirmed (debug)', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, groupId: groupOrTeamId, userNames: userName }) });
    assert(postSpy.notCalled);
  });

  it('removes the specified owner from owners and members endpoint of the Microsoft 365 Group specified by id with accepted prompt', async () => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/owners/${userId}/$ref`) {
        return;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/members/${userId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: userName }) });
    assert(memberDeleteCallIssued);
  });

  it('removes the specified owner from owners and members endpoint of the specified Microsoft 365 Group when prompt confirmed', async () => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/owners/${userId}/$ref`) {
        return;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/members/${userId}/$ref`) {
        return;
      }

      throw 'Invalid request';

    });

    await command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: userName, force: true }) });
    assert(memberDeleteCallIssued);
  });

  it('removes the specified member from members endpoint of the specified Microsoft 365 Group when prompt confirmed', async () => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/owners/${userId}/$ref`) {
        throw {
          "response": {
            "status": 404
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/members/${userId}/$ref`) {
        return;
      }

      throw 'Invalid request';

    });

    await command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: userName, force: true }) });
    assert(memberDeleteCallIssued);
  });

  it('removes the specified members of the specified Microsoft 365 Group specified by teamName', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupOrTeamName).resolves(groupOrTeamId);

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/owners/${userId}/$ref`) {
        return;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/members/${userId}/$ref`) {
        throw {
          "response": {
            "status": 404
          }
        };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ teamName: groupOrTeamName, userNames: userName, verbose: true }) });
    assert(deleteStub.calledTwice);
  });

  it('removes the specified members specified by ids of the specified Microsoft 365 Team specified by teamId', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupOrTeamName).resolves(groupOrTeamId);

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/owners/${userId}/$ref`) {
        return;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/members/${userId}/$ref`) {
        throw {
          "response": {
            "status": 404
          }
        };
      }

      throw 'Invalid request';

    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ teamId: groupOrTeamId, ids: userId, verbose: true }) });
    assert(deleteStub.calledTwice);
  });

  it('removes the specified members of the specified Microsoft 365 Group specified by groupName', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupOrTeamName).resolves(groupOrTeamId);

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/owners/${userId}/$ref`) {
        return;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/members/${userId}/$ref`) {
        return;
      }

      throw 'Invalid request';

    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ teamName: groupOrTeamName, userNames: userName }) });
    assert(deleteStub.calledTwice);
  });

  it('does not fail if the user is not owner or member of the specified Microsoft 365 Group when prompt confirmed', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/owners/${userId}/$ref`) {
        return {
          "response": {
            "status": 404
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/members/${userId}/$ref`) {
        throw {
          "response": {
            "status": 404
          }
        };
      }

      throw 'Invalid request';

    });


    await command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: userName, force: true }) });
    assert(deleteStub.calledTwice);
  });

  it('stops removal if an unknown error message is thrown when deleting the owner', async () => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      memberDeleteCallIssued = true;

      // for example... you must have at least one owner
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/owners/${userId}/$ref`) {
        return {
          "response": {
            "status": 400
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/members/${userId}/$ref`) {
        throw {
          "response": {
            "status": 404
          }
        };
      }

      throw 'Invalid request';

    });

    await command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: userName, force: true }) });
    assert(memberDeleteCallIssued);
  });

  it('correctly retrieves user but does not find the Group Microsoft 365 group', async () => {
    const errorMessage = `Resource '${groupOrTeamId}' does not exist or one of its queried reference-property objects are not present.`;

    sinonUtil.restore(cli.promptForConfirmation);
    sinonUtil.restore(entraGroup.isUnifiedGroup);

    sinon.stub(entraGroup, 'isUnifiedGroup').rejects(
      {
        error: {
          code: 'Request_ResourceNotFound',
          message: errorMessage,
          innerError: {
            date: '2024-09-16T22:06:30',
            'request-id': 'c43610b0-70c0-4c00-8c40-ff26b5f37f00',
            'client-request-id': 'c43610b0-70c0-4c00-8c40-ff26b5f37f00'
          }
        }
      }
    );
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: userName }) }),
      new CommandError(errorMessage));
  });

  it('correctly retrieves user and handle error removing owner from specified Microsoft 365 group', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/owners/${userId}/$ref`) {
        throw {
          response: {
            status: 400,
            data: {
              error: { 'odata.error': { message: { value: 'Invalid object identifier' } } }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: userName }) } as any),
      new CommandError('Invalid object identifier'));
  });

  it('correctly retrieves user and handle error removing member from specified Microsoft 365 group', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/owners/${userId}/$ref`) {
        return {
          err: {
            response: {
              status: 404
            }
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupOrTeamId}/members/${userId}/$ref`) {
        throw {
          response: {
            status: 400,
            data: {
              error: { 'odata.error': { message: { value: 'Invalid object identifier' } } }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: userName }) }),
      new CommandError('Invalid object identifier'));
  });

  it('throws error when the group is not a unified group', async () => {
    sinonUtil.restore(entraGroup.isUnifiedGroup);
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ groupId: groupOrTeamId, userNames: 'anne.matthews@contoso.onmicrosoft.com', force: true }) }),
      new CommandError(`Specified group with id '${groupOrTeamId}' is not a Microsoft 365 group.`));
  });
});
