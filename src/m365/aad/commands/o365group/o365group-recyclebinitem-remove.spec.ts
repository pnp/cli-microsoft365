import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./o365group-recyclebinitem-remove');

describe(commands.O365GROUP_RECYCLEBINITEM_REMOVE, () => {
  const validGroupId = '00000000-0000-0000-0000-000000000000';
  const validGroupDisplayName = 'Dev Team';
  const validGroupMailNickname = 'Devteam';

  const singleGroupsResponse = {
    value: [
      {
        id: validGroupId,
        displayName: validGroupDisplayName,
        mailNickname: validGroupDisplayName,
        mail: 'Devteam@contoso.com',
        groupTypes: [
          "Unified"
        ]
      }
    ]
  };

  const multipleGroupsResponse = {
    value: [
      {
        id: validGroupId,
        displayName: validGroupDisplayName,
        mailNickname: validGroupDisplayName,
        mail: 'Devteam@contoso.com',
        groupTypes: [
          "Unified"
        ]
      },
      {
        id: validGroupId,
        displayName: validGroupDisplayName,
        mailNickname: validGroupDisplayName,
        mail: 'Devteam@contoso.com',
        groupTypes: [
          "Unified"
        ]
      }
    ]
  };

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
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.O365GROUP_RECYCLEBINITEM_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when id is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        id: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({
      options: {
        id: validGroupId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified group when confirm option not passed with id', async () => {
    await command.action(logger, {
      options: {
        id: validGroupId
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified group when confirm option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: {
        id: validGroupId
      }
    });
    assert(deleteSpy.notCalled);
  });

  it('throws error message when no group was found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        mailNickname: validGroupMailNickname,
        confirm: true
      }
    }), new CommandError(`The specified group '${validGroupMailNickname}' does not exist.`));
  });

  it('throws error message when multiple groups were found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'`) {
        return multipleGroupsResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        mailNickname: validGroupMailNickname,
        confirm: true
      }
    }), new CommandError(`Multiple groups with name '${validGroupMailNickname}' found: ${multipleGroupsResponse.value.map(x => x.id).join(',')}.`));
  });

  it('correctly deletes group by id with confirm flag', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: validGroupId,
        confirm: true
      }
    });
  });

  it('correctly deletes group by id', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        id: validGroupId
      }
    });
  });

  it('correctly deletes group by displayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupDisplayName)}'`) {
        return singleGroupsResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        displayName: validGroupDisplayName
      }
    });
  });

  it('correctly deletes group by mailNickname', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'`) {
        return singleGroupsResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        mailNickname: validGroupMailNickname
      }
    });
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'An error has occurred'
      }
    };
    sinon.stub(request, 'delete').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        id: validGroupId,
        confirm: true
      }
    }), new CommandError("An error has occurred"));
  });
}); 
