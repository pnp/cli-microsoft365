import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import { RoleDefinition } from '../roledefinition/RoleDefinition.js';
import command from './list-roleassignment-add.js';
import { settingsNames } from '../../../../settingsNames.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

describe(commands.LIST_ROLEASSIGNMENT_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const webUrl = 'https://contoso.sharepoint.com/sites/Marketing';

  const userResponse = {
    Id: 11,
    IsHiddenInUI: false,
    LoginName: 'i:0#.f|membership|john.doe@contoso.com',
    Title: 'John Doe',
    PrincipalType: 1,
    Email: 'john.doe@contoso.com',
    Expiration: '',
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: false,
    UserId: {
      NameId: '10032002473c5ae3',
      NameIdIssuer: 'urn:federation:microsoftonline'
    },
    UserPrincipalName: 'john.doe@contoso.com'
  };

  const entraGroupResponse = {
    Id: 15,
    IsHiddenInUI: false,
    LoginName: 'c:0o.c|federateddirectoryclaimprovider|27ae47f1-48f1-46f3-980b-d3c1470e398d',
    Title: 'Marketing members',
    PrincipalType: 1,
    Email: '',
    Expiration: '',
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: false,
    UserId: null,
    UserPrincipalName: null
  };

  const groupResponse = {
    Id: 12,
    IsHiddenInUI: false,
    LoginName: "groupname",
    Title: "groupname",
    PrincipalType: 8,
    AllowMembersEditMembership: false,
    AllowRequestToJoinLeave: false,
    AutoAcceptRequestToJoinLeave: false,
    Description: "",
    OnlyAllowMembersViewMembership: true,
    OwnerTitle: "John Doe",
    RequestToJoinLeaveEmailSetting: null
  };

  const roledefinitionResponse: RoleDefinition = {
    BasePermissions: {
      High: 176,
      Low: 138612833
    },
    Description: "Can view pages and list items and download documents.",
    Hidden: false,
    Id: 1073741827,
    Name: "Read",
    Order: 128,
    RoleTypeKind: 2,
    BasePermissionsValue: [
      "ViewListItems",
      "OpenItems",
      "ViewVersions",
      "ViewFormPages",
      "Open",
      "ViewPages",
      "CreateSSCSite",
      "BrowseUserInfo",
      "UseClientIntegration",
      "UseRemoteAPIs",
      "CreateAlerts"
    ],
    RoleTypeKindValue: "Reader"
  };

  const graphGroup = {
    id: '27ae47f1-48f1-46f3-980b-d3c1470e398d',
    deletedDateTime: null,
    classification: null,
    createdDateTime: '2024-03-22T20:18:37Z',
    creationOptions: [],
    description: null,
    displayName: 'Marketing',
    expirationDateTime: null,
    groupTypes: [
      'Unified'
    ],
    isAssignableToRole: null,
    mail: 'Marketing@milanhdev.onmicrosoft.com',
    mailEnabled: true,
    mailNickname: 'Marketing',
    membershipRule: null,
    membershipRuleProcessingState: null,
    onPremisesDomainName: null,
    onPremisesLastSyncDateTime: null,
    onPremisesNetBiosName: null,
    onPremisesSamAccountName: null,
    onPremisesSecurityIdentifier: null,
    onPremisesSyncEnabled: null,
    preferredDataLocation: null,
    preferredLanguage: null,
    proxyAddresses: [
      'SPO:SPO_de7704ba-415d-4dd0-9bbd-fa565007a87e@SPO_18c58817-3bc9-489d-ac63-f7264fb357e5',
      'SMTP:Marketing@milanhdev.onmicrosoft.com'
    ],
    renewedDateTime: '2024-03-22T20:18:37Z',
    resourceBehaviorOptions: [],
    resourceProvisioningOptions: [],
    securityEnabled: true,
    securityIdentifier: 'S-1-12-1-665733105-1190349041-3268610968-2369326662',
    theme: null,
    uniqueName: null,
    visibility: 'Private',
    onPremisesProvisioningErrors: [],
    serviceProvisioningErrors: []
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => settingName === settingsNames.prompt ? false : defaultValue);
    sinon.stub(entraGroup, 'getGroupById').withArgs(graphGroup.id).resolves(graphGroup);
    sinon.stub(entraGroup, 'getGroupByDisplayName').withArgs(graphGroup.displayName).resolves(graphGroup);
    sinon.stub(spo, 'ensureEntraGroup').withArgs(webUrl, graphGroup).resolves(entraGroupResponse);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      spo.getGroupByName,
      spo.ensureUser,
      spo.getRoleDefinitionByName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_ROLEASSIGNMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 'abc', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the principalId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the roleDefinitionId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the roleDefinitionId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the upn option is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', upn: 'invalid', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the upn option is a valid user principal name', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', upn: 'john.doe@contoso.com', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the entraGroupId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', entraGroupId: 'invalid', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the entraGroupId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', entraGroupId: '37455d5c-e466-4e49-8eba-808b5acec21b', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds role assignment on list by title and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('test')/roleassignments/addroleassignment(principalid='11',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('adds role assignment on list by id and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/addroleassignment(principalid='11',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('adds role assignment on list by url and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetList('%2Fsites%2Fdocuments')/roleassignments/addroleassignment(principalid='11',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listUrl: 'sites/documents',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('adds role assignment on list and gets principal id by upn', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/addroleassignment(principalid='11',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'ensureUser').resolves(userResponse);

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/addroleassignment(principalid='11',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'User cannot be found.';
    sinon.stub(spo, 'ensureUser').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    }), new CommandError(error));
  });

  it('adds role assignment on list and gets principal id by group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/addroleassignment(principalid='12',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getGroupByName').resolves(groupResponse);

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when group does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/addroleassignment(principalid='11',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'Group cannot be found';
    sinon.stub(spo, 'getGroupByName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    }), new CommandError(error));
  });

  it('adds role assignment on list and gets role definition id by role definition name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/addroleassignment(principalid='11',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getRoleDefinitionByName').resolves(roledefinitionResponse);

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    });
  });

  it('correctly handles error when role definition does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/addroleassignment(principalid='11',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'No roledefinition is found for Read';
    sinon.stub(spo, 'getRoleDefinitionByName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    }), new CommandError(error));
  });

  it('correctly adds role assignments for a Microsoft Entra group specified by ID', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/addroleassignment(principalid='15',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/Marketing',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        entraGroupId: graphGroup.id,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly adds role assignments for a Microsoft Entra group specified by display name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/addroleassignment(principalid='15',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/Marketing',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        entraGroupName: graphGroup.displayName,
        roleDefinitionId: 1073741827
      }
    });
  });
});
