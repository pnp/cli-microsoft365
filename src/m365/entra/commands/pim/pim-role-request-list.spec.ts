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
import command from './pim-role-request-list.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { CommandError } from '../../../../Command.js';

describe(commands.PIM_ROLE_REQUEST_LIST, () => {
  const userId = '61b0c52f-a902-4769-9a09-c6628335b00a';
  const userName = 'john.doe@contoso.com';
  const groupId = '1caf7dcd-7e83-4c3a-94f7-932a1299c844';
  const groupName = 'SharePoint Administrators';
  const createdDateTime = '2024-01-01T12:00:00Z';
  const status = 'Revoked';

  const unifiedRoleAssignmentScheduleRequestResponse = [
    {
      "id": "80231d2f-95a1-47a5-8339-acf3d71efec7",
      "status": "Revoked",
      "createdDateTime": "2024-02-12T14:08:38.82Z",
      "completedDateTime": null,
      "approvalId": null,
      "customData": null,
      "action": "adminRemove",
      "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
      "roleDefinitionId": "f28a1f50-f6e7-4571-818b-6a12f2af6b6c",
      "directoryScopeId": "/",
      "appScopeId": null,
      "isValidationOnly": false,
      "targetScheduleId": null,
      "justification": null,
      "scheduleInfo": null,
      "createdBy": {
        "application": null,
        "device": null,
        "user": {
          "displayName": null,
          "id": "893f9116-e024-4bc6-8e98-54c245129485"
        }
      },
      "ticketInfo": {
        "ticketNumber": null,
        "ticketSystem": null
      },
      "roleDefinition": {
        "displayName": "SharePoint Administrator"
      }
    },
    {
      "id": "da1cb564-78ba-4198-b94a-613d892ed73e",
      "status": "Granted",
      "createdDateTime": "2023-10-15T12:35:08.167Z",
      "completedDateTime": null,
      "approvalId": null,
      "customData": null,
      "action": "adminAssign",
      "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "directoryScopeId": "/administrativeUnits/0a22c83d-c4ac-43e2-bb5e-87af3015d49f",
      "appScopeId": null,
      "isValidationOnly": false,
      "targetScheduleId": null,
      "justification": null,
      "scheduleInfo": {
        "startDateTime": "2024-02-05T07:11:09.0773878Z",
        "recurrence": null,
        "expiration": {
          "type": "noExpiration",
          "endDateTime": null,
          "duration": null
        }
      },
      "createdBy": {
        "application": null,
        "device": null,
        "user": {
          "displayName": null,
          "id": "893f9116-e024-4bc6-8e98-54c245129485"
        }
      },
      "ticketInfo": {
        "ticketNumber": null,
        "ticketSystem": null
      },
      "roleDefinition": {
        "displayName": "User Administrator"
      }
    }
  ];

  const unifiedRoleAssignmentScheduleRequestTransformedResponse = [
    {
      "id": "80231d2f-95a1-47a5-8339-acf3d71efec7",
      "status": "Revoked",
      "createdDateTime": "2024-02-12T14:08:38.82Z",
      "completedDateTime": null,
      "approvalId": null,
      "customData": null,
      "action": "adminRemove",
      "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
      "roleDefinitionId": "f28a1f50-f6e7-4571-818b-6a12f2af6b6c",
      "directoryScopeId": "/",
      "appScopeId": null,
      "isValidationOnly": false,
      "targetScheduleId": null,
      "justification": null,
      "scheduleInfo": null,
      "createdBy": {
        "application": null,
        "device": null,
        "user": {
          "displayName": null,
          "id": "893f9116-e024-4bc6-8e98-54c245129485"
        }
      },
      "ticketInfo": {
        "ticketNumber": null,
        "ticketSystem": null
      },
      "roleDefinitionName": "SharePoint Administrator"
    },
    {
      "id": "da1cb564-78ba-4198-b94a-613d892ed73e",
      "status": "Granted",
      "createdDateTime": "2023-10-15T12:35:08.167Z",
      "completedDateTime": null,
      "approvalId": null,
      "customData": null,
      "action": "adminAssign",
      "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "directoryScopeId": "/administrativeUnits/0a22c83d-c4ac-43e2-bb5e-87af3015d49f",
      "appScopeId": null,
      "isValidationOnly": false,
      "targetScheduleId": null,
      "justification": null,
      "scheduleInfo": {
        "startDateTime": "2024-02-05T07:11:09.0773878Z",
        "recurrence": null,
        "expiration": {
          "type": "noExpiration",
          "endDateTime": null,
          "duration": null
        }
      },
      "createdBy": {
        "application": null,
        "device": null,
        "user": {
          "displayName": null,
          "id": "893f9116-e024-4bc6-8e98-54c245129485"
        }
      },
      "ticketInfo": {
        "ticketNumber": null,
        "ticketSystem": null
      },
      "roleDefinitionName": "User Administrator"
    }
  ];

  const unifiedRoleAssignmentScheduleRequestWithPrincipalResponse = [
    {
      "id": "da1cb564-78ba-4198-b94a-613d892ed73e",
      "status": "Granted",
      "createdDateTime": "2023-10-15T12:35:08.167Z",
      "completedDateTime": null,
      "approvalId": null,
      "customData": null,
      "action": "adminAssign",
      "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "directoryScopeId": "/administrativeUnits/0a22c83d-c4ac-43e2-bb5e-87af3015d49f",
      "appScopeId": null,
      "isValidationOnly": false,
      "targetScheduleId": null,
      "justification": null,
      "scheduleInfo": {
        "startDateTime": "2024-02-05T07:11:09.0773878Z",
        "recurrence": null,
        "expiration": {
          "type": "noExpiration",
          "endDateTime": null,
          "duration": null
        }
      },
      "createdBy": {
        "application": null,
        "device": null,
        "user": {
          "displayName": null,
          "id": "893f9116-e024-4bc6-8e98-54c245129485"
        }
      },
      "ticketInfo": {
        "ticketNumber": null,
        "ticketSystem": null
      },
      "roleDefinition": {
        "displayName": "User Administrator"
      },
      "principal": {
        "id": "61b0c52f-a902-4769-9a09-c6628335b00a",
        "displayName": "John Doe",
        "userPrincipalName": "JohnDoe@contoso.onmicrosoft.com",
        "mail": "JohnDoe@contoso.onmicrosoft.com",
        "businessPhones": [
          "+1 425 555 0109"
        ],
        "givenName": "John",
        "jobTitle": "Retail Manager",
        "mobilePhone": null,
        "officeLocation": "18/2111",
        "preferredLanguage": "en-US",
        "surname": "Doe"
      }
    }
  ];

  const unifiedRoleAssignmentScheduleRequestWithPrincipalTransformedResponse = [
    {
      "id": "da1cb564-78ba-4198-b94a-613d892ed73e",
      "status": "Granted",
      "createdDateTime": "2023-10-15T12:35:08.167Z",
      "completedDateTime": null,
      "approvalId": null,
      "customData": null,
      "action": "adminAssign",
      "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "directoryScopeId": "/administrativeUnits/0a22c83d-c4ac-43e2-bb5e-87af3015d49f",
      "appScopeId": null,
      "isValidationOnly": false,
      "targetScheduleId": null,
      "justification": null,
      "scheduleInfo": {
        "startDateTime": "2024-02-05T07:11:09.0773878Z",
        "recurrence": null,
        "expiration": {
          "type": "noExpiration",
          "endDateTime": null,
          "duration": null
        }
      },
      "createdBy": {
        "application": null,
        "device": null,
        "user": {
          "displayName": null,
          "id": "893f9116-e024-4bc6-8e98-54c245129485"
        }
      },
      "ticketInfo": {
        "ticketNumber": null,
        "ticketSystem": null
      },
      "roleDefinitionName": "User Administrator",
      "principal": {
        "id": "61b0c52f-a902-4769-9a09-c6628335b00a",
        "displayName": "John Doe",
        "userPrincipalName": "JohnDoe@contoso.onmicrosoft.com",
        "mail": "JohnDoe@contoso.onmicrosoft.com",
        "businessPhones": [
          "+1 425 555 0109"
        ],
        "givenName": "John",
        "jobTitle": "Retail Manager",
        "mobilePhone": null,
        "officeLocation": "18/2111",
        "preferredLanguage": "en-US",
        "surname": "Doe"
      }
    }
  ];

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
      request.get,
      entraUser.getUserIdByUpn,
      entraGroup.getGroupIdByDisplayName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PIM_ROLE_REQUEST_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when userName is a valid user principal name', async () => {
    const actual = await command.validate({ options: { userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: groupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when createdDateTime is a valid ISO 8601 date', async () => {
    const actual = await command.validate({ options: { createdDateTime: '2024-02-20T08:00:00Z' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when status is set to one of allowed values', async () => {
    const actual = await command.validate({ options: { status: 'Granted' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userName is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { userName: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when createdDateTime is not a valid ISO 8601 date', async () => {
    const actual = await command.validate({ options: { createdDateTime: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when status has invalid value', async () => {
    const actual = await command.validate({ options: { status: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should get a list of PIM requests', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$expand=roleDefinition($select=displayName)') {
        return {
          value: unifiedRoleAssignmentScheduleRequestResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });

    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScheduleRequestTransformedResponse));
  });

  it('should get a list of PIM requests for a user specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$filter=principalId eq '${userId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleRequestResponse[0]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId, verbose: true } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleRequestTransformedResponse[0]]));
  });

  it('should get a list of PIM requests for a user specified by name', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$filter=principalId eq '${userId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleRequestResponse[0]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, verbose: true } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleRequestTransformedResponse[0]]));
  });

  it('should get a list of PIM requests for a group specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$filter=principalId eq '${groupId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleRequestResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, verbose: true } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleRequestTransformedResponse[1]]));
  });

  it('should get a list of PIM requests for a group specified by name', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$filter=principalId eq '${groupId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleRequestResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: groupName, verbose: true } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleRequestTransformedResponse[1]]));
  });

  it('should get a list of PIM requests from specified start date', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$filter=createdDateTime ge ${createdDateTime}&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleRequestResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { createdDateTime: createdDateTime } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleRequestTransformedResponse[1]]));
  });

  it('should get a list of PIM requests with specified status', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$filter=status eq '${status}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleRequestResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { status: status } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleRequestTransformedResponse[1]]));
  });

  it('should get a list of PIM requests for a user specified by id from specified start date and with specified status', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$filter=principalId eq '${userId}' and createdDateTime ge ${createdDateTime} and status eq '${status}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleRequestResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId, createdDateTime: createdDateTime, status: status } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleRequestTransformedResponse[1]]));
  });

  it('should get a list of PIM requests with details about principal that were assigned', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$expand=roleDefinition($select=displayName),principal') {
        return {
          value: unifiedRoleAssignmentScheduleRequestWithPrincipalResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { includePrincipalDetails: true } });

    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScheduleRequestWithPrincipalTransformedResponse));
  });

  it('should get a list of PIM requests for a group specified by name from specified start date with details about principal that were assigned', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$filter=principalId eq '${groupId}' and createdDateTime ge ${createdDateTime}&$expand=roleDefinition($select=displayName),principal`) {
        return {
          value: unifiedRoleAssignmentScheduleRequestWithPrincipalResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: groupName, createdDateTime: createdDateTime, includePrincipalDetails: true } });

    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScheduleRequestWithPrincipalTransformedResponse));
  });

  it('handles error when retrieving PIM requests failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests?$expand=roleDefinition($select=displayName)') {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred')
    );
  });
});