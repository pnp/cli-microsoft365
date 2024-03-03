import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import command from './group-remove.js';
import { settingsNames } from '../../../../settingsNames.js';
import { formatting } from '../../../../utils/formatting.js';
import aadCommands from '../../aadCommands.js';

describe(commands.GROUP_REMOVE, () => {
  const groupId = '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a';
  const displayName = 'CLI Test Group';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
      request.get,
      request.delete,
      entraGroup.getGroupIdByDisplayName,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_REMOVE);
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
    assert.deepStrictEqual(alias, [aadCommands.GROUP_REMOVE]);
  });

  it('removes the specified group by id without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, id: groupId, force: true } });
    assert(deleteRequestStub.called);
  });

  it('removes the specified group by displayName when passing the force option', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(groupId);

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, displayName: displayName, force: true } });
    assert(deleteRequestStub.called);
    assert(confirmationStub.notCalled);
  });

  it('removes the specified group by displayName while prompting for confirmation', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(groupId);

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, displayName: displayName } });
    assert(deleteRequestStub.called);
    assert(confirmationStub.calledOnce);
  });

  it('throws an error when group by id cannot be found', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const error = {
      error: {
        code: 'Request_ResourceNotFound',
        message: `Resource '${groupId}' does not exist or one of its queried reference-property objects are not present.`,
        innerError: {
          date: '2023-08-30T14:32:41',
          'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
          'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
        }
      }
    };

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, id: groupId, force: true } }),
      new CommandError(error.error.message));
  });

  it('prompts before removing the specified group when confirm option not passed', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { id: groupId } });

    assert(confirmationStub.calledOnce);
  });

  it('handles error when multiple groups with the specified displayName found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
        return {
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(request, 'delete').rejects('DELETE request executed');

    await assert.rejects(command.action(logger, {
      options: {
        displayName: displayName,
        force: true
      }
    }), new CommandError(`Multiple groups with name 'CLI Test Group' found. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`));
  });

  it('handles selecting single result when multiple groups with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
        return {
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/9b1b1e42-794b-4c71-93ac-5ed92488b67f`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: displayName, force: true } });
    assert(deleteRequestStub.called);
  });

  it('aborts removing group when prompt not confirmed', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    const deleteSpy = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { id: groupId } });
    assert(deleteSpy.notCalled);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: groupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
