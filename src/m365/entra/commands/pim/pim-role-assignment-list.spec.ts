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
import command from './pim-role-assignment-list.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { CommandError } from '../../../../Command.js';

describe(commands.PIM_ROLE_ASSIGNMENT_LIST, () => {
  const userId = '61b0c52f-a902-4769-9a09-c6628335b00a';
  const userName = 'john.doe@contoso.com';
  const groupId = '1caf7dcd-7e83-4c3a-94f7-932a1299c844';
  const groupName = 'SharePoint Administrators';
  const startDateTime = '2024-01-01T12:00:00Z';

  const unifiedRoleAssignmentScheduleInstanceResponse = [
    {
      "id": "5wuT_mJe20eRr5jDpJo4sS_FsGECqWlHmgnGYoM1sApj5okazO8RSY336VRxQAXe-2",
      "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "directoryScopeId": "/administrativeUnits/1a89e663-efcc-4911-8df7-e954714005de",
      "appScopeId": null,
      "startDateTime": "2023-11-15T12:24:32.773Z",
      "endDateTime": null,
      "assignmentType": "Assigned",
      "memberType": "Direct",
      "roleAssignmentOriginId": "5wuT_mJe20eRr5jDpJo4sS_FsGECqWlHmgnGYoM1sApj5okazO8RSY336VRxQAXe-2",
      "roleAssignmentScheduleId": "36bd668f-3a40-455f-a40a-64074fde4a18",
      "roleDefinition": {
        "displayName": "User Administrator"
      }
    },
    {
      "id": "5wuT_mJe20eRr5jDpJo4seCabNh9bS9BgvTNJIBCEKw-1",
      "principalId": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "directoryScopeId": "/",
      "appScopeId": null,
      "startDateTime": "2024-02-12T08:47:02.91Z",
      "endDateTime": null,
      "assignmentType": "Assigned",
      "memberType": "Direct",
      "roleAssignmentOriginId": "5wuT_mJe20eRr5jDpJo4seCabNh9bS9BgvTNJIBCEKw-1",
      "roleAssignmentScheduleId": "5f2c16a0-4212-4fa2-afae-fc8bfdc527b6",
      "roleDefinition": {
        "displayName": "SharePoint Administrator"
      }
    }
  ];

  const unifiedRoleAssignmentScheduleInstanceWithPrincipalResponse = [
    {
      "id": "5wuT_mJe20eRr5jDpJo4seCabNh9bS9BgvTNJIBCEKw-1",
      "principalId": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "directoryScopeId": "/",
      "appScopeId": null,
      "startDateTime": "2024-02-12T08:47:02.91Z",
      "endDateTime": null,
      "assignmentType": "Assigned",
      "memberType": "Direct",
      "roleAssignmentOriginId": "5wuT_mJe20eRr5jDpJo4seCabNh9bS9BgvTNJIBCEKw-1",
      "roleAssignmentScheduleId": "5f2c16a0-4212-4fa2-afae-fc8bfdc527b6",
      "roleDefinition": {
        "displayName": "SharePoint Administrator"
      },
      "principal": {
        "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
        "displayName": "SharePoint Administrators",
        "mail": "SharePointAdministrators@contoso.com",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2024-02-12T08:45:59Z",
        "creationOptions": [],
        "description": null,
        "expirationDateTime": null,
        "groupTypes": [
          "Unified"
        ],
        "isAssignableToRole": true,
        "mailEnabled": true,
        "mailNickname": "SharePointAdministrators",
        "membershipRule": null,
        "membershipRuleProcessingState": null,
        "onPremisesDomainName": null,
        "onPremisesLastSyncDateTime": null,
        "onPremisesNetBiosName": null,
        "onPremisesSamAccountName": null,
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "preferredLanguage": null,
        "proxyAddresses": [
          "SMTP:SharePointAdministrators@contoso.com"
        ],
        "renewedDateTime": "2024-02-12T08:45:59Z",
        "resourceBehaviorOptions": [],
        "resourceProvisioningOptions": [],
        "securityEnabled": true,
        "securityIdentifier": "S-1-12-1-1234567890-1234567890-123456789-1234567890",
        "theme": null,
        "visibility": "Private",
        "onPremisesProvisioningErrors": [],
        "serviceProvisioningErrors": []
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
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      cli.promptForSelection,
      entraUser.getUserIdByUpn,
      entraGroup.getGroupIdByDisplayName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PIM_ROLE_ASSIGNMENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: groupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when startDateTime is a valid ISO 8601 date', async () => {
    const actual = await command.validate({ options: { startDateTime: '2024-02-20T08:00:00Z' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when startDateTime is not a valid ISO 8601 date', async () => {
    const actual = await command.validate({ options: { startDateTime: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should get a list of role assignments', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$expand=roleDefinition($select=displayName)`) {
        return {
          value: unifiedRoleAssignmentScheduleInstanceResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });

    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScheduleInstanceResponse));
  });

  it('should get a list of role assignments for a user specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$filter=principalId eq '${userId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleInstanceResponse[0]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleInstanceResponse[0]]));
  });

  it('should get a list of role assignments for a user specified by name', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$filter=principalId eq '${userId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleInstanceResponse[0]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, verbose: true } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleInstanceResponse[0]]));
  });

  it('should get a list of role assignments for a group specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$filter=principalId eq '${groupId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleInstanceResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleInstanceResponse[1]]));
  });

  it('should get a list of role assignments for a group specified by name', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$filter=principalId eq '${groupId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleInstanceResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: groupName, verbose: true } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleInstanceResponse[1]]));
  });

  it('should get a list of role assignments from specified start date', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$filter=startDateTime ge ${startDateTime}&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleInstanceResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { startDateTime: startDateTime } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleInstanceResponse[1]]));
  });

  it('should get a list of role assignments for a user specified by id from specified start date', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$filter=principalId eq '${userId}' and startDateTime ge ${startDateTime}&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentScheduleInstanceResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId, startDateTime: startDateTime } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentScheduleInstanceResponse[1]]));
  });

  it(`correctly shows deprecation warning for option 'includePrincipalDetails'`, async () => {
    const chalk = (await import('chalk')).default;
    const loggerErrSpy = sinon.spy(logger, 'logToStderr');

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$expand=roleDefinition($select=displayName),principal`) {
        return {
          value: unifiedRoleAssignmentScheduleInstanceWithPrincipalResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { includePrincipalDetails: true } });
    assert(loggerErrSpy.calledWith(chalk.yellow(`Parameter 'includePrincipalDetails' is deprecated. Please use 'withPrincipalDetails' instead`)));

    sinonUtil.restore(loggerErrSpy);
  });

  it('should get a list of role assignments with details about principal that were assigned', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$expand=roleDefinition($select=displayName),principal`) {
        return {
          value: unifiedRoleAssignmentScheduleInstanceWithPrincipalResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { withPrincipalDetails: true } });

    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScheduleInstanceWithPrincipalResponse));
  });

  it('should get a list of role assignments for a group specified by name from specified start date with details about principal that were assigned', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$filter=principalId eq '${groupId}' and startDateTime ge ${startDateTime}&$expand=roleDefinition($select=displayName),principal`) {
        return {
          value: unifiedRoleAssignmentScheduleInstanceWithPrincipalResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: groupName, startDateTime: startDateTime, withPrincipalDetails: true } });

    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScheduleInstanceWithPrincipalResponse));
  });

  it('handles error when retrieving role assignments failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?$expand=roleDefinition($select=displayName)`) {
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