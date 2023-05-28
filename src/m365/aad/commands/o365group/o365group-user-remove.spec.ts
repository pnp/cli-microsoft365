import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import teamsCommands from '../../../teams/commands';
import commands from '../../commands';
const command: Command = require('./o365group-user-remove');

describe(commands.O365GROUP_USER_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      global.setTimeout,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.O365GROUP_USER_REMOVE);
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
    const actual = await command.validate({
      options: {
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both groupId and teamId are specified', async () => {
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

  it('prompts before removing the specified user from the specified Microsoft 365 Group when confirm option not passed', async () => {
    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('prompts before removing the specified user from the specified Team when confirm option not passed (debug)', async () => {
    await command.action(logger, { options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified user from the specified Microsoft 365 Group when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } });
    assert(postSpy.notCalled);
  });

  it('aborts removing the specified user from the specified Microsoft 365 Group when confirm option not passed and prompt not confirmed (debug)', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

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

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

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

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } });
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

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } });
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

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } });
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


    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } });
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

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } });
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

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
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

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

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

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

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

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await assert.rejects(command.action(logger, { options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } } as any), new CommandError("Invalid request"));
  });
});
