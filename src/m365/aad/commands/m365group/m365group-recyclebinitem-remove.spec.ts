import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './m365group-recyclebinitem-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.M365GROUP_RECYCLEBINITEM_REMOVE, () => {
  let cli: Cli;
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
  let promptIssued: boolean = false;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
    sinon.stub(Cli.getInstance(), 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
      Cli.promptForConfirmation,
      Cli.handleMultipleResultsFound,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_RECYCLEBINITEM_REMOVE);
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

  it('prompts before removing the specified group when force option not passed with id', async () => {
    await command.action(logger, {
      options: {
        id: validGroupId
      }
    });

    assert(promptIssued);
  });

  it('aborts removing the specified group when force option not passed and prompt not confirmed', async () => {
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
        force: true
      }
    }), new CommandError(`The specified group '${validGroupMailNickname}' does not exist.`));
  });

  it('throws error message when multiple groups were found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'`) {
        return multipleGroupsResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        mailNickname: validGroupMailNickname,
        force: true
      }
    }), new CommandError("Multiple groups with name 'Devteam' found. Found: 00000000-0000-0000-0000-000000000000."));
  });

  it('handles selecting single result when multiple groups with the specified name found and cli is set to prompt', async () => {
    let removeRequestIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupDisplayName)}'`) {
        return multipleGroupsResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validGroupId}`) {
        removeRequestIssued = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves(singleGroupsResponse.value[0]);

    sinonUtil.restore(Cli.promptForConfirmation);

    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        displayName: validGroupDisplayName
      }
    });
    assert(removeRequestIssued);
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
        force: true
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
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

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
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

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
        force: true
      }
    }), new CommandError("An error has occurred"));
  });
}); 
