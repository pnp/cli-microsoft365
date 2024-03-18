import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import commands from '../../commands.js';
import command from './user-groupmembership-list.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { entraUser } from '../../../../utils/entraUser.js';

describe(commands.USER_GROUPMEMBERSHIP_LIST, () => {
  const userId = "68be84bf-a585-4776-80b3-30aa5207aa21";
  const userName = "john.doe@contoso.onmicrosoft.com";
  const groupMembershipResponse: any = {
    value: [
      "2f64f70d-386b-489f-805a-670cad739fde",
      "ff0554cc-8aa8-40f2-a369-ed604503fb79",
      "0a0bf25a-2de0-40de-9908-c96941a2615b"
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      entraUser.getUserIdByEmail,
      entraUser.getUserIdByUpn
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_GROUPMEMBERSHIP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation if userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if userName is a valid UPN', async () => {
    const actual = await command.validate({ options: { userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if userEmail is a valid email', async () => {
    const actual = await command.validate({ options: { userEmail: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { userName: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('retrieves groups memberships for a user specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/getMemberGroups` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "securityEnabledOnly": false
        })) {
        return groupMembershipResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId } });
    assert(loggerLogSpy.calledOnceWithExactly(groupMembershipResponse.value));
  });

  it('retrieves groups memberships for a user specified by UPN', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/getMemberGroups` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "securityEnabledOnly": false
        })) {
        return groupMembershipResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName } });
    assert(loggerLogSpy.calledOnceWithExactly(groupMembershipResponse.value));
  });

  it('retrieves groups memberships for a user specified by email', async () => {
    sinon.stub(entraUser, 'getUserIdByEmail').withArgs(userName).resolves(userId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/getMemberGroups` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "securityEnabledOnly": false
        })) {
        return groupMembershipResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userEmail: userName } });
    assert(loggerLogSpy.calledOnceWithExactly(groupMembershipResponse.value));
  });

  it('retrieves memberships in security enabled groups only for a user specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/getMemberGroups` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "securityEnabledOnly": true
        })) {
        return groupMembershipResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId, securityEnabledOnly: true } });
    assert(loggerLogSpy.calledOnceWithExactly(groupMembershipResponse.value));
  });

  it('fails validation if userEmail is not a valid email', async () => {
    const actual = await command.validate({ options: { userEmail: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'post').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: { userId: userId } }), new CommandError(errorMessage));
  });
});