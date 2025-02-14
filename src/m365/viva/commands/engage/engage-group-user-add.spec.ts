import assert from 'assert';
import sinon from 'sinon';
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
import command from './engage-group-user-add.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.ENGAGE_GROUP_USER_ADD, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
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
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_GROUP_USER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      throw {
        "error": {
          "base": "An error has occurred."
        }
      };
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { groupId: 10123123 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('groupId must be a number', async () => {
    const actual = await command.validate({ options: { groupId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('id must be a number', async () => {
    const actual = await command.validate({ options: { groupId: 10, id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('calls the service if the current user is added to the group', async () => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { debug: true, groupId: 1231231 } });

    assert(requestPostedStub.called);
  });

  it('calls the service if the user 989998789 is added to the group 1231231', async () => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, groupId: 1231231, id: 989998789 } });

    assert(requestPostedStub.called);
  });

  it('calls the service if the user suzy@contoso.com is added to the group 1231231', async () => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, groupId: 1231231, email: "suzy@contoso.com" } });

    assert(requestPostedStub.called);
  });
}); 
