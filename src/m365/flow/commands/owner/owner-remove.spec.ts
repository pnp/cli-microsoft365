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

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      aadGroup.getGroupByDisplayName,
      aadUser.getUserIdByUpn,
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
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

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, name: flowName, userId: userId } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBodyUser);
  });

  it('deletes owner from flow by userName', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').callsFake(async () => {
      return userId;
    });
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === requestUrlNoAdmin) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, name: flowName, userName: userName } });
    assert.deepStrictEqual(JSON.stringify(postStub.lastCall.args[0].data), JSON.stringify(requestBodyUser));
  });

  it('deletes owner from flow by groupId as admin', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === requestUrlAdmin) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, name: flowName, groupId: groupId, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBodyGroup);
  });

  it('deletes owner from flow by groupName as admin', async () => {
    sinon.stub(aadGroup, 'getGroupByDisplayName').callsFake(async () => {
      return { id: groupId };
    });
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === requestUrlAdmin) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, name: flowName, groupName: groupName, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBodyGroup);
  });

  it('throws error when no environment found', async () => {
    const error = {
      'error': {
        'code': 'EnvironmentAccessDenied',
        'message': `Access to the environment '${environmentName}' is denied.`
      }
    };
    sinon.stub(request, 'post').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName, name: flowName, userId: userId } } as any),
      new CommandError(error.error.message));
  });

  it('throws error when Flow not found', async () => {
    const error = {
      'error': {
        'code': 'FlowNotFound',
        'message': `Could not find flow '${flowName}'.`
      }
    };
    sinon.stub(request, 'post').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName, name: flowName, userId: userId } } as any),
      new CommandError(error.error.message));
  });

  it('fails validation if flowName is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, name: 'invalid', userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, name: flowName, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, name: flowName, groupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if username is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, name: flowName, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if groupName passed', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, name: flowName, groupName: groupName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
