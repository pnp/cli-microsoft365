import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { aadUser } from '../../../../utils/aadUser';
import { aadGroup } from '../../../../utils/aadGroup';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
import { formatting } from '../../../../utils/formatting';
import { session } from '../../../../utils/session';
const command: Command = require('./owner-remove');

describe(commands.OWNER_REMOVE, () => {
  const environmentName = 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6';
  const flowName = '1c6ee23a-a835-44bc-a4f5-462b658efc12';
  const userId = 'bb46fcd8-2104-4f3d-982c-eb74677251c2';
  const userName = 'john@contoso.com';
  const groupId = '37a0264d-fea4-4e87-8e5e-e574ff878cf2';
  const groupName = 'Test Group';
  const requestUrlNoAdmin = `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(environmentName)}/flows/${formatting.encodeQueryParameter(flowName)}/modifyPermissions?api-version=2016-11-01`;
  const requestUrlAdmin = `https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}/flows/${formatting.encodeQueryParameter(flowName)}/modifyPermissions?api-version=2016-11-01`;
  const requestBodyUser = { 'delete': [{ 'id': userId }] };
  const requestBodyGroup = { 'delete': [{ 'id': groupId }] };

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
      aadGroup.getGroupIdByDisplayName,
      aadUser.getUserIdByUpn,
      Cli.prompt,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.OWNER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('deletes owner from flow by userId', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === requestUrlNoAdmin) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, flowName: flowName, userId: userId, confirm: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBodyUser);
  });

  it('deletes owner from flow by userName', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').resolves(userId);
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === requestUrlNoAdmin) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, flowName: flowName, userName: userName, confirm: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBodyUser);
  });

  it('deletes owner from flow by groupId as admin when prompt confirmed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === requestUrlAdmin) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, flowName: flowName, groupId: groupId, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBodyGroup);
  });

  it('deletes owner from flow by groupName as admin', async () => {
    sinon.stub(aadGroup, 'getGroupIdByDisplayName').resolves(groupId);
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === requestUrlAdmin) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, flowName: flowName, groupName: groupName, asAdmin: true, confirm: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBodyGroup);
  });

  it('throws error when no environment found', async () => {
    const error = {
      error: {
        code: 'EnvironmentAccessDenied',
        message: `Access to the environment '${environmentName}' is denied.`
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName, flowName: flowName, userId: userId, confirm: true } } as any),
      new CommandError(error.error.message));
  });

  it('prompts before removing the specified owner from a flow when confirm option not passed', async () => {
    await command.action(logger, { options: { environmentName: environmentName, flowName: flowName, useName: userName } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified owner from a flow when option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    await command.action(logger, { options: { environmentName: environmentName, flowName: flowName, useName: userName } });
    assert(postSpy.notCalled);
  });

  it('fails validation if flowName is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: 'invalid', userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, groupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if username is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if groupName passed', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, groupName: groupName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
