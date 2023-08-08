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
import command from './o365group-recyclebinitem-restore.js';

describe(commands.O365GROUP_RECYCLEBINITEM_RESTORE, () => {
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
      request.post,
      global.setTimeout
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.O365GROUP_RECYCLEBINITEM_RESTORE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('restores the specified group by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deleteditems/${validGroupId}/restore`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        id: validGroupId
      }
    });
  });

  it('correctly restores group by displayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupDisplayName)}'`) {
        return singleGroupsResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deleteditems/${validGroupId}/restore`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        displayName: validGroupDisplayName
      }
    });
  });

  it('correctly restores group by mailNickname', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'`) {
        return singleGroupsResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deleteditems/${validGroupId}/restore`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        mailNickname: validGroupMailNickname
      }
    });
  });

  it('correctly handles error when group is not found', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'Group Not Found.' } } } });

    await assert.rejects(command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } } as any),
      new CommandError('Group Not Found.'));
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
    }), new CommandError(`Multiple groups with name '${validGroupMailNickname}' found: ${multipleGroupsResponse.value.map(x => x.id).join(',')}.`));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
