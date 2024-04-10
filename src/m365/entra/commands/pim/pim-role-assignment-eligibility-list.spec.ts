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
import command from './pim-role-assignment-eligibility-list.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { CommandError } from '../../../../Command.js';

describe(commands.PIM_ROLE_ASSIGNMENT_ELIGIBILITY_LIST, () => {
  const userId = '61b0c52f-a902-4769-9a09-c6628335b00a';
  const userName = 'john.doe@contoso.com';
  const groupId = '1caf7dcd-7e83-4c3a-94f7-932a1299c844';
  const groupName = 'SharePoint Administrators';

  const unifiedRoleAssignmentEligibilityScheduleInstanceResponse = [
    {
      "id": "XrtkCdube02sKVjnlIYqQBht8lJR0U9DrhSkqDEisrI-1-e",
      "principalId": "52f26d18-d151-434f-ae14-a4a83122b2b2",
      "roleDefinitionId": "0964bb5e-9bdb-4d7b-ac29-58e794862a40",
      "directoryScopeId": "/",
      "appScopeId": null,
      "startDateTime": "2024-04-08T10:14:01.153Z",
      "endDateTime": null,
      "memberType": "Direct",
      "roleEligibilityScheduleId": "7a135e3d-5be5-403c-bdad-47ccbac434e3",
      "roleDefinition": {
        "displayName": "Search Administrator"
      }
    },
    {
      "id": "YMROdH45rUKkYos_l0egLC_FsGECqWlHmgnGYoM1sAo-1-e",
      "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
      "roleDefinitionId": "744ec460-397e-42ad-a462-8b3f9747a02c",
      "directoryScopeId": "/",
      "appScopeId": null,
      "startDateTime": "2024-04-08T10:13:04.913Z",
      "endDateTime": "2025-04-08T10:12:36.9Z",
      "memberType": "Direct",
      "roleEligibilityScheduleId": "0606b8a1-ba92-42b7-804c-8e32dfdec2b8",
      "roleDefinition": {
        "displayName": "Knowledge Manager"
      }
    }
  ];

  const unifiedRoleAssignmentEligibilityScheduleInstanceWithPrincipalResponse = [
    {
      "id": "XrtkCdube02sKVjnlIYqQBht8lJR0U9DrhSkqDEisrI-1-e",
      "principalId": "52f26d18-d151-434f-ae14-a4a83122b2b2",
      "roleDefinitionId": "0964bb5e-9bdb-4d7b-ac29-58e794862a40",
      "directoryScopeId": "/",
      "appScopeId": null,
      "startDateTime": "2024-04-08T10:14:01.153Z",
      "endDateTime": null,
      "memberType": "Direct",
      "roleEligibilityScheduleId": "7a135e3d-5be5-403c-bdad-47ccbac434e3",
      "roleDefinition": {
        "displayName": "Search Administrator"
      },
      "principal": {
        "id": "52f26d18-d151-434f-ae14-a4a83122b2b2",
        "displayName": "Alex Wilber",
        "userPrincipalName": "AlexW@contoso.onmicrosoft.com",
        "mail": "AlexW@contoso.onmicrosoft.com",
        "businessPhones": [
          "+1 858 555 0110"
        ],
        "givenName": "Alex",
        "jobTitle": "Marketing Assistant",
        "mobilePhone": null,
        "officeLocation": "131/1104",
        "preferredLanguage": "en-US",
        "surname": "Wilber"
      }
    },
    {
      "id": "YMROdH45rUKkYos_l0egLC_FsGECqWlHmgnGYoM1sAo-1-e",
      "principalId": "61b0c52f-a902-4769-9a09-c6628335b00a",
      "roleDefinitionId": "744ec460-397e-42ad-a462-8b3f9747a02c",
      "directoryScopeId": "/",
      "appScopeId": null,
      "startDateTime": "2024-04-08T10:13:04.913Z",
      "endDateTime": "2025-04-08T10:12:36.9Z",
      "memberType": "Direct",
      "roleEligibilityScheduleId": "0606b8a1-ba92-42b7-804c-8e32dfdec2b8",
      "roleDefinition": {
        "displayName": "Knowledge Manager"
      },
      "principal": {
        "id": "61b0c52f-a902-4769-9a09-c6628335b00a",
        "displayName": "Adele Vance",
        "userPrincipalName": "AdeleV@contoso.onmicrosoft.com",
        "mail": "AdeleV@contoso.onmicrosoft.com",
        "businessPhones": [
          "+1 425 555 0109"
        ],
        "givenName": "Adele",
        "jobTitle": "Retail Manager",
        "mobilePhone": null,
        "officeLocation": "18/2111",
        "preferredLanguage": "en-US",
        "surname": "Vance"
      }
    }
  ];

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
    assert.strictEqual(command.name, commands.PIM_ROLE_ASSIGNMENT_ELIGIBILITY_LIST);
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

  it('fails validation when userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should get a list of eligible roles for any user', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$expand=roleDefinition($select=displayName)`) {
        return {
          value: unifiedRoleAssignmentEligibilityScheduleInstanceResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });

    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentEligibilityScheduleInstanceResponse));
  });

  it('should get a list of eligible roles for a user specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$filter=principalId eq '${userId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentEligibilityScheduleInstanceResponse[0]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentEligibilityScheduleInstanceResponse[0]]));
  });

  it('should get a list of eligible roles for a user specified by name', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$filter=principalId eq '${userId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentEligibilityScheduleInstanceResponse[0]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, verbose: true } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentEligibilityScheduleInstanceResponse[0]]));
  });

  it('should get a list of eligible roles for a group specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$filter=principalId eq '${groupId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentEligibilityScheduleInstanceResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentEligibilityScheduleInstanceResponse[1]]));
  });

  it('should get a list of eligible roles for a group specified by name', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$filter=principalId eq '${groupId}'&$expand=roleDefinition($select=displayName)`) {
        return {
          value: [
            unifiedRoleAssignmentEligibilityScheduleInstanceResponse[1]
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: groupName, verbose: true } });

    assert(loggerLogSpy.calledOnceWithExactly([unifiedRoleAssignmentEligibilityScheduleInstanceResponse[1]]));
  });

  it('should get a list of eligible roles with details about principals that were assigned', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$expand=roleDefinition($select=displayName),principal`) {
        return {
          value: unifiedRoleAssignmentEligibilityScheduleInstanceWithPrincipalResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { includePrincipalDetails: true } });

    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentEligibilityScheduleInstanceWithPrincipalResponse));
  });

  it('handles error when retrieving a list of eligible roles failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$expand=roleDefinition($select=displayName)`) {
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