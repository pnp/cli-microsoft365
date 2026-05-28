import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './owner-ensure.js';

describe(commands.OWNER_ENSURE, () => {
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
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      entraGroup.getGroupByDisplayName,
      entraUser.getUserIdByUpn,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.OWNER_ENSURE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if flowName is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironmentName, flowName: 'invalid', userId: validUserId, roleName: validRoleName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironmentName, flowName: validFlowName, userId: 'invalid', roleName: validRoleName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if groupId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironmentName, flowName: validFlowName, groupId: 'invalid', roleName: validRoleName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if username is not a valid user principal name', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironmentName, flowName: validFlowName, userName: 'invalid', roleName: validRoleName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if roleName is not a valid role name', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironmentName, flowName: validFlowName, userName: validUserName, roleName: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when no owner identifier is provided', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironmentName, flowName: validFlowName, roleName: validRoleName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when multiple owner identifiers are provided', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironmentName, flowName: validFlowName, userId: validUserId, groupId: validGroupId, roleName: validRoleName });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when required parameters are provided', () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironmentName, flowName: validFlowName, userName: validUserName, roleName: validRoleName });
    assert.strictEqual(actual.success, true);
  });

  it('adds owner to a flow with userId', async () => {
    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validUserId,
              type: 'User'
            },
            roleName: 'CanView'
          }
        }
      ]
    };

    const postRequestStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(validEnvironmentName)}/flows/${formatting.encodeQueryParameter(validFlowName)}/modifyPermissions?api-version=2016-11-01`) {
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
              id: validUserId,
              type: 'User'
            },
            roleName: 'CanEdit'
          }
        }
      ]
    };

    sinon.stub(entraUser, 'getUserIdByUpn').resolves(validUserId);

    const postRequestStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(validEnvironmentName)}/flows/${formatting.encodeQueryParameter(validFlowName)}/modifyPermissions?api-version=2016-11-01`) {
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
              id: validGroupId,
              type: 'Group'
            },
            roleName: 'CanEdit'
          }
        }
      ]
    };

    const postRequestStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(validEnvironmentName)}/flows/${formatting.encodeQueryParameter(validFlowName)}/modifyPermissions?api-version=2016-11-01`) {
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
              id: validGroupId,
              type: 'Group'
            },
            roleName: 'CanEdit'
          }
        }
      ]
    };

    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(validGroupId);

    const postRequestStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(validEnvironmentName)}/flows/${formatting.encodeQueryParameter(validFlowName)}/modifyPermissions?api-version=2016-11-01`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironmentName, flowName: validFlowName, groupName: validGroupName, roleName: validRoleName, asAdmin: true } });
    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('handles selecting single result when multiple groups with the specified name found and cli is set to prompt', async () => {
    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validGroupId,
              type: 'Group'
            },
            roleName: 'CanEdit'
          }
        }
      ]
    };

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'&$select=id`) {
        return {
          value: [
            { id: validGroupId },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: validGroupId });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(validEnvironmentName)}/flows/${formatting.encodeQueryParameter(validFlowName)}/modifyPermissions?api-version=2016-11-01`) {
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
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, { options: { environmentName: validEnvironmentName, flowName: validFlowName, roleName: validRoleName, userId: validUserId } } as any),
      new CommandError(error.error.message));
  });
});