import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { odata } from '../../../../utils/odata';
import { formatting } from '../../../../utils/formatting';
const command: Command = require('./group-recyclebinitem-remove');

describe(commands.GROUP_RECYCLEBINITEM_REMOVE, () => {
  const groupId = '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a';
  const displayName = 'CLI Test Group';
  const groupResponse = [
    {
      id: groupId
    }
  ];


  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      odata.getAllItems,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_RECYCLEBINITEM_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the specified group from recyclebin by id without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, id: groupId, force: true } });
    assert(deleteRequestStub.called);
  });

  it('removes the specified group from recyclebin by displayName while prompting for confirmation', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
        return groupResponse;
      }
      throw 'Invalid request';
    });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { verbose: true, displayName: displayName } });
    assert(deleteRequestStub.called);
  });

  it('throws an error when group from recyclebin by displayname returns multiple results', async () => {
    const secondGroupId = 'bf1ee687-ee55-4ddb-9fec-f3409672cedd';
    const groupResponseClone = [...groupResponse];
    groupResponseClone.push({ id: secondGroupId });

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
        return groupResponseClone;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, displayName: displayName, force: true } }),
      new CommandError(`Multiple deleted groups with name 'CLI Test Group' found: ${groupId},${secondGroupId}.`));
  });

  it('throws an error when group from recyclebin by displayname returns no results', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
        return [];
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, displayName: displayName, force: true } }),
      new CommandError(`The specified deleted group '${displayName}' does not exist.`));
  });

  it('throws an error when group from recyclebin by id cannot be found', async () => {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, id: groupId, force: true } }),
      new CommandError(error.error.message));
  });

  it('prompts before removing the specified group setting when confirm option not passed', async () => {
    await command.action(logger, { options: { id: groupId } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the group setting when prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');

    await command.action(logger, { options: { id: groupId } });
    assert(deleteSpy.notCalled);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: groupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});