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
const command: Command = require('./owner-add');

describe(commands.OWNER_ADD, () => {
  const validEnvironmentName = 'Default-6a2903af-9c03-4c02-a50b-e7419599925b';
  const validFlowName = '784670e6-199a-4993-ae13-4b6747a0cd5d';
  const validUserId = 'd2481133-e3ed-4add-836d-6e200969dd03';
  const validUserName = 'john.doe@contoso.com';
  const validGroupId = 'c6c4b4e0-cd72-4d64-8ec2-cfbd0388ec16';
  const validGroupName = 'CLI Group';
  const validRoleName = 'CanEdit';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.OWNER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if flowName is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, flowName: 'invalid', userId: validUserId, roleName: validRoleName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, flowName: validFlowName, userId: 'invalid', roleName: validRoleName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, flowName: validFlowName, groupId: 'invalid', roleName: validRoleName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if username is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, flowName: validFlowName, userName: 'invalid', roleName: validRoleName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if roleName is not a valid role name', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, flowName: validFlowName, userName: validUserName, roleName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when required parameters are provided', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironmentName, flowName: validFlowName, userName: validUserName, roleName: validRoleName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds owner to the flow with userId', async () => {
    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validUserId
            },
            roleName: 'CanView',
            type: 'User'
          }
        }
      ]
    };

    const postRequestStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(validEnvironmentName)}/flows/${formatting.encodeQueryParameter(validFlowName)}/modifyPermissions?api-version=2016-11-01`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironmentName, flowName: validFlowName, userId: validUserId, roleName: 'CanView' } });
    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('adds owner to the flow with userName', async () => {
    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validUserId
            },
            roleName: 'CanEdit',
            type: 'User'
          }
        }
      ]
    };

    sinon.stub(aadUser, 'getUserIdByUpn').callsFake(async () => {
      return validUserId;
    });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(validEnvironmentName)}/flows/${formatting.encodeQueryParameter(validFlowName)}/modifyPermissions?api-version=2016-11-01`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironmentName, flowName: validFlowName, userName: validUserName, roleName: validRoleName } });
    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('adds owner to the flow with groupId as admin', async () => {
    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validGroupId
            },
            roleName: 'CanEdit',
            type: 'Group'
          }
        }
      ]
    };

    const postRequestStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(validEnvironmentName)}/flows/${formatting.encodeQueryParameter(validFlowName)}/modifyPermissions?api-version=2016-11-01`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironmentName, flowName: validFlowName, groupId: validGroupId, roleName: validRoleName, asAdmin: true } });
    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('adds owner to the flow with groupName as admin', async () => {
    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validGroupId
            },
            roleName: 'CanEdit',
            type: 'Group'
          }
        }
      ]
    };

    sinon.stub(aadGroup, 'getGroupByDisplayName').callsFake(async () => {
      return { id: validGroupId };
    });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(validEnvironmentName)}/flows/${formatting.encodeQueryParameter(validFlowName)}/modifyPermissions?api-version=2016-11-01`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironmentName, flowName: validFlowName, groupName: validGroupName, roleName: validRoleName, asAdmin: true } });
    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('correctly handles API OData error', async () => {
    const error = {
      error: {
        message: 'Could not find flow'
      }
    };
    sinon.stub(request, 'post').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, { options: { environmentName: validEnvironmentName, flowName: validFlowName, roleName: validRoleName, userId: validUserId } } as any),
      new CommandError(error.error.message));
  });
});