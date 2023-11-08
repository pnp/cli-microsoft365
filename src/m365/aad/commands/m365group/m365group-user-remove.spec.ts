import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import teamsCommands from '../../../teams/commands.js';
import commands from '../../commands.js';
import command from './m365group-user-remove.js';
import { settingsNames } from '../../../../settingsNames.js';
import { aadGroup } from '../../../../utils/aadGroup.js';

describe(commands.M365GROUP_USER_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(aadGroup, 'isUnifiedGroup').resolves(true);
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      global.setTimeout,
      Cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_USER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.strictEqual((alias && alias.indexOf(teamsCommands.USER_REMOVE) > -1), true);
  });

  it('fails validation if the groupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        groupId: 'not-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'not-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither the groupId nor the teamID are provided.', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both groupId and teamId are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid groupId and userName are specified', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified user from the specified Microsoft 365 Group when force option not passed', async () => {
    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } });

    assert(promptIssued);
  });

  it('prompts before removing the specified user from the specified Team when force option not passed (debug)', async () => {
    await command.action(logger, { options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } });

    assert(promptIssued);
  });

  it('aborts removing the specified user from the specified Microsoft 365 Group when force option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } });
    assert(postSpy.notCalled);
  });

  it('aborts removing the specified user from the specified Microsoft 365 Group when force option not passed and prompt not confirmed (debug)', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } });
    assert(postSpy.notCalled);
  });

  it('removes the specified owner from owners and members endpoint of the specified Microsoft 365 Group with accepted prompt', async () => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return {
          "value": "00000000-0000-0000-0000-000000000001"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return {
          response: {
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          }
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000" }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000" }]
        };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } });
    assert(memberDeleteCallIssued);
  });

  it('removes the specified owner from owners and members endpoint of the specified Microsoft 365 Group when prompt confirmed', async () => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return {
          "value": "00000000-0000-0000-0000-000000000001"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return {
          response: {
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          }
        };
      }

      throw 'Invalid request';
    });


    sinon.stub(request, 'delete').callsFake(async (opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000" }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000" }]
        };
      }

      throw 'Invalid request';

    });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", force: true } });
    assert(memberDeleteCallIssued);
  });

  it('removes the specified member from members endpoint of the specified Microsoft 365 Group when prompt confirmed', async () => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return {
          "value": "00000000-0000-0000-0000-000000000001"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return {
          response: {
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          }
        };
      }

      throw 'Invalid request';
    });


    sinon.stub(request, 'delete').callsFake(async (opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        throw {
          "response": {
            "status": 404
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000" }]
        };
      }

      throw 'Invalid request';

    });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", force: true } });
    assert(memberDeleteCallIssued);
  });

  it('removes the specified owners from owners endpoint of the specified Microsoft 365 Group when prompt confirmed', async () => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return {
          "value": "00000000-0000-0000-0000-000000000001"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return {
          response: {
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          }
        };
      }

      throw 'Invalid request';
    });


    sinon.stub(request, 'delete').callsFake(async (opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000" }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        throw {
          "response": {
            "status": 404
          }
        };
      }

      throw 'Invalid request';

    });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", force: true } });
    assert(memberDeleteCallIssued);
  });

  it('does not fail if the user is not owner or member of the specified Microsoft 365 Group when prompt confirmed', async () => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return {
          "value": "00000000-0000-0000-0000-000000000001"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return {
          response: {
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          }
        };
      }

      throw 'Invalid request';
    });


    sinon.stub(request, 'delete').callsFake(async (opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return {
          "response": {
            "status": 404
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        throw {
          "response": {
            "status": 404
          }
        };
      }

      throw 'Invalid request';

    });


    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", force: true } });
    assert(memberDeleteCallIssued);
  });

  it('stops removal if an unknown error message is thrown when deleting the owner', async () => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return {
          "value": "00000000-0000-0000-0000-000000000001"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return {
          response: {
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          }
        };
      }

      throw 'Invalid request';
    });


    sinon.stub(request, 'delete').callsFake(async (opts) => {
      memberDeleteCallIssued = true;

      // for example... you must have at least one owner
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return {
          "response": {
            "status": 400
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        throw {
          "response": {
            "status": 404
          }
        };
      }

      throw 'Invalid request';

    });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", force: true } });
    assert(memberDeleteCallIssued);
  });

  it('correctly retrieves user but does not find the Group Microsoft 365 group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return {
          "value": "00000000-0000-0000-0000-000000000001"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        throw "Invalid object identifier";
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
    await assert.rejects(command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } } as any),
      new CommandError('Invalid object identifier'));
  });

  it('correctly retrieves user and handle error removing owner from specified Microsoft 365 group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return {
          "value": "00000000-0000-0000-0000-000000000001"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return {
          response: {
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          }
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
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

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } } as any),
      new CommandError('Invalid object identifier'));
  });

  it('correctly retrieves user and handle error removing member from specified Microsoft 365 group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return {
          "value": "00000000-0000-0000-0000-000000000001"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return {
          response: {
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          }
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return {
          "err": {
            "response": {
              "status": 404
            }
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
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

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } } as any),
      new CommandError('Invalid object identifier'));
  });

  it('correctly skips execution when specified user is not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews.not.found%40contoso.onmicrosoft.com/id`) {
        throw "Resource 'anne.matthews.not.found%40contoso.onmicrosoft.com' does not exist or one of its queried reference-property objects are not present.";
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'delete').resolves();

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, { options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } } as any), new CommandError("Invalid request"));
  });

  it('throws error when the group is not a unified group', async () => {
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    sinonUtil.restore(aadGroup.isUnifiedGroup);
    sinon.stub(aadGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { groupId: groupId, userName: 'anne.matthews@contoso.onmicrosoft.com', force: true } } as any),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });
});
