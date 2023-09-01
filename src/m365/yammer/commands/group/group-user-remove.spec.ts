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
import commands from '../../commands.js';
import command from './group-user-remove.js';

describe(commands.GROUP_USER_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let requests: any[];
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
    requests = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_USER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'delete').callsFake(async () => {
      throw {
        "error": {
          "base": "An error has occurred."
        }
      };
    });

    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

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

  it('calls the service if the current user is removed from the group', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { debug: true, groupId: 1231231 } });

    assert(requestDeleteStub.called);
  });

  it('calls the service if the user 989998789 is removed from the group 1231231 with the confirm command', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, groupId: 1231231, id: 989998789, force: true } });

    assert(requestDeleteStub.called);
  });

  it('calls the service if the user 989998789 is removed from the group 1231231', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { debug: true, groupId: 1231231, id: 989998789 } });

    assert(requestDeleteStub.called);
  });

  it('prompts before removal when confirmation argument not passed', async () => {
    const promptStub: sinon.SinonStub = sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));

    await command.action(logger, { options: { groupId: 1231231, id: 989998789 } });

    assert(promptStub.called);
  });

  it('aborts execution when prompt not confirmed', async () => {
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));

    await command.action(logger, { options: { groupId: 1231231, id: 989998789 } });

    assert(requests.length === 0);
  });
}); 
