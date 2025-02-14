import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import command from './pim-role-assignment-remove.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { CommandError } from '../../../../Command.js';

describe(commands.PIM_ROLE_ASSIGNMENT_REMOVE, () => {
  const roleDefinitionId = 'f1417aa3-bf0b-4cc5-a845-a0b2cf11f690';
  const roleDefinitionName = 'SharePoint Administrator';
  const userId = '61b0c52f-a902-4769-9a09-c6628335b00a';
  const userName = 'john.doe@contoso.com';
  const groupId = '1caf7dcd-7e83-4c3a-94f7-932a1299c844';
  const groupName = 'SharePoint Administrators';
  const roleAssignmentResponseTenantScope = {
    "id": "3b74089b-5078-441f-9ebe-4eada46fe826",
    "status": "Revoked",
    "createdDateTime": "2024-02-11T19:27:41.0546404Z",
    "completedDateTime": null,
    "approvalId": null,
    "customData": null,
    "action": "adminRemove",
    "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
    "roleDefinitionId": "f28a1f50-f6e7-4571-818b-6a12f2af6b6c",
    "directoryScopeId": "/",
    "appScopeId": null,
    "isValidationOnly": false,
    "targetScheduleId": "3b74089b-5078-441f-9ebe-4eada46fe826",
    "justification": "Remove SharePoint Administrator role",
    "createdBy": {
      "application": null,
      "device": null,
      "user": {
        "displayName": null,
        "id": "893f9116-e024-4bc6-8e98-54c245129485"
      }
    },
    "scheduleInfo": null,
    "ticketInfo": {
      "ticketNumber": null,
      "ticketSystem": null
    }
  };

  const roleAssignmentResponseAdminUnitScope = {
    "id": "39c06837-9692-42c1-838d-cd4d53247df6",
    "status": "Revoked",
    "createdDateTime": "2024-02-11T19:37:06.7494657Z",
    "completedDateTime": null,
    "approvalId": null,
    "customData": null,
    "action": "adminRemove",
    "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
    "roleDefinitionId": "f28a1f50-f6e7-4571-818b-6a12f2af6b6c",
    "directoryScopeId": "/administrativeUnits/81bb36e4-f4c6-4984-8e56-d4f8feae9e09",
    "appScopeId": null,
    "isValidationOnly": false,
    "targetScheduleId": "39c06837-9692-42c1-838d-cd4d53247df6",
    "justification": "Remove SharePoint Administrator role for admin unit",
    "createdBy": {
      "application": null,
      "device": null,
      "user": {
        "displayName": null,
        "id": "893f9116-e024-4bc6-8e98-54c245129485"
      }
    },
    "scheduleInfo": null,
    "ticketInfo": {
      "ticketNumber": null,
      "ticketSystem": null
    }
  };

  const roleAssignmentResponseApplicationScope = {
    "id": "6d2ca8e1-2230-42a5-80c3-2d0febc814cf",
    "status": "Revoked",
    "createdDateTime": "2024-02-12T08:33:02.1822857Z",
    "completedDateTime": null,
    "approvalId": null,
    "customData": null,
    "action": "adminRemove",
    "principalId": "3d284fb2-1895-4eb6-9289-2dcc7475b6db",
    "roleDefinitionId": "9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3",
    "directoryScopeId": "/94446d35-4df6-45da-a17f-c601310a8342",
    "appScopeId": null,
    "isValidationOnly": false,
    "targetScheduleId": "6d2ca8e1-2230-42a5-80c3-2d0febc814cf",
    "justification": "Remove Application Administrator role for group for two days",
    "createdBy": {
      "application": null,
      "device": null,
      "user": {
        "displayName": null,
        "id": "893f9116-e024-4bc6-8e98-54c245129485"
      }
    },
    "scheduleInfo": null,
    "ticketInfo": {
      "ticketNumber": null,
      "ticketSystem": null
    }
  };

  const roleAssignmentResponseWithTicketInfo = {
    "id": "5f2c16a0-4212-4fa2-afae-fc8bfdc527b6",
    "status": "Revoked",
    "createdDateTime": "2024-02-12T08:47:01.8016121Z",
    "completedDateTime": null,
    "approvalId": null,
    "customData": null,
    "action": "adminRemove",
    "principalId": "d86c9ae0-6d7d-412f-82f4-cd24804210ac",
    "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
    "directoryScopeId": "/",
    "appScopeId": null,
    "isValidationOnly": false,
    "targetScheduleId": "5f2c16a0-4212-4fa2-afae-fc8bfdc527b6",
    "justification": "Remove User Administrator role for group, ticket details included",
    "createdBy": {
      "application": null,
      "device": null,
      "user": {
        "displayName": null,
        "id": "893f9116-e024-4bc6-8e98-54c245129485"
      }
    },
    "scheduleInfo": {
      "startDateTime": "2024-02-12T08:47:02.7244107Z",
      "recurrence": null,
      "expiration": {
        "type": "noExpiration",
        "endDateTime": null,
        "duration": null
      }
    },
    "ticketInfo": {
      "ticketNumber": "MSFT-2024",
      "ticketSystem": "JIRA"
    }
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      cli.promptForSelection,
      roleDefinition.getRoleDefinitionByDisplayName,
      entraUser.getUserIdByUpn,
      entraGroup.getGroupIdByDisplayName,
      accessToken.isAppOnlyAccessToken,
      accessToken.getUserIdFromAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PIM_ROLE_ASSIGNMENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when roleDefinitionId is a valid GUID', async () => {
    const actual = await command.validate({ options: { roleDefinitionId: roleDefinitionId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: userId, roleDefinitionName: 'Global Administrator' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: groupId, roleDefinitionName: 'Global Administrator' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when administrativeUnitId is a valid GUID', async () => {
    const actual = await command.validate({ options: { roleDefinitionId: roleDefinitionId, administrativeUnitId: '81bb36e4-f4c6-4984-8e56-d4f8feae9e09' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when applicationId is a valid GUID', async () => {
    const actual = await command.validate({ options: { roleDefinitionId: roleDefinitionId, applicationId: '94446d35-4df6-45da-a17f-c601310a8342' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when roleDefinitionId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { roleDefinitionId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'foo', roleDefinitionName: 'Global Administrator' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: 'foo', roleDefinitionName: 'Global Administrator' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when administrativeUnitId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { roleDefinitionId: roleDefinitionId, administrativeUnitId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when applicationId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { roleDefinitionId: roleDefinitionId, applicationId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly requests deactivation of role specified by id for a user specified by id tenant-wide scope', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "principalId": userId,
          "roleDefinitionId": roleDefinitionId,
          "directoryScopeId": "/",
          "action": "adminRemove",
          "justification": "Remove user from SharePoint Administrator role",
          "ticketInfo": {
          }
        })) {
        return roleAssignmentResponseTenantScope;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        userId: userId,
        justification: 'Remove user from SharePoint Administrator role'
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(roleAssignmentResponseTenantScope));
  });

  it('correctly requests deactivation of role specified by name for a user specified by name with administrative unit scope', async () => {
    sinon.stub(roleDefinition, 'getRoleDefinitionByDisplayName').withArgs(roleDefinitionName).resolves({ id: roleDefinitionId, displayName: roleDefinitionName });
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "principalId": userId,
          "roleDefinitionId": roleDefinitionId,
          "directoryScopeId": "/administrativeUnits/81bb36e4-f4c6-4984-8e56-d4f8feae9e09",
          "action": "adminRemove",
          "justification": "Remove user from SharePoint Administrator role for admin unit",
          "ticketInfo": {
          }
        })) {
        return roleAssignmentResponseAdminUnitScope;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        roleDefinitionName: roleDefinitionName,
        userName: userName,
        administrativeUnitId: '81bb36e4-f4c6-4984-8e56-d4f8feae9e09',
        justification: 'Remove user from SharePoint Administrator role for admin unit',
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(roleAssignmentResponseAdminUnitScope));
  });

  it('correctly requests deactivation of role specified by id for a group specified by id with application scope', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "principalId": groupId,
          "roleDefinitionId": roleDefinitionId,
          "directoryScopeId": "/94446d35-4df6-45da-a17f-c601310a8342",
          "action": "adminRemove",
          "justification": "Remove Application Administrator role for group",
          "ticketInfo": {
          }
        })) {
        return roleAssignmentResponseApplicationScope;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        groupId: groupId,
        applicationId: '94446d35-4df6-45da-a17f-c601310a8342',
        justification: 'Remove Application Administrator role for group'
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(roleAssignmentResponseApplicationScope));
  });

  it('correctly requests deactivation of role specified by id for a group specified by name with ticket details', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "principalId": groupId,
          "roleDefinitionId": roleDefinitionId,
          "directoryScopeId": "/",
          "action": "adminRemove",
          "justification": "Remove User Administrator role for group, ticket details included",
          "ticketInfo": {
            "ticketNumber": "MSFT-2024",
            "ticketSystem": "JIRA"
          }
        })) {
        return roleAssignmentResponseWithTicketInfo;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        groupName: groupName,
        justification: 'Remove User Administrator role for group, ticket details included',
        ticketSystem: 'JIRA',
        ticketNumber: 'MSFT-2024',
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(roleAssignmentResponseWithTicketInfo));
  });

  it('correctly requests deactivation of role specified by id for a current user', async () => {
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: '',
      accessToken: ''
    };
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(userId);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "principalId": userId,
          "roleDefinitionId": roleDefinitionId,
          "directoryScopeId": "/",
          "action": "selfDeactivate",
          "justification": "Remove SharePoint Administrator role",
          "ticketInfo": {
          }
        })) {
        return roleAssignmentResponseTenantScope;
      }

      throw opts.data;
    });

    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        justification: 'Remove SharePoint Administrator role',
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(roleAssignmentResponseTenantScope));
  });

  it('fails deactivation of role specified by id for a current user when running as app', async () => {
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: '',
      accessToken: ''
    };
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { roleDefinitionId: roleDefinitionId, verbose: true } }), new CommandError('When running with application permissions either userId, userName, groupId or groupName is required'));
  });

  it('throws an error during self deactivation when role assignment does not exist', async () => {
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: '',
      accessToken: ''
    };
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(userId);
    const error = {
      error: {
        code: 'RoleAssignmentDoesNotExist',
        message: 'The Role assignment does not exist.',
        innerError: {
          date: '2023-10-27T12:24:36',
          'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
          'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
        }
      }
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "principalId": userId,
          "roleDefinitionId": roleDefinitionId,
          "directoryScopeId": "/",
          "action": "selfDeactivate",
          "justification": "Remove SharePoint Administrator role",
          "ticketInfo": {
          }
        })) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { roleDefinitionId: roleDefinitionId, justification: 'Remove SharePoint Administrator role' } }), new CommandError(error.error.message));
  });
});