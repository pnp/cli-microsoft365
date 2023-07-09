import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { spo } from '../../../../utils/spo';
import { RoleDefinition } from '../roledefinition/RoleDefinition';
const command: Command = require('./list-roleassignment-add');

describe(commands.LIST_ROLEASSIGNMENT_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
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

  const groupResponse = {
    Id: 11,
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      spo.getGroupByName,
      spo.getUserByEmail,
      spo.getRoleDefinitionByName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
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

  it('add role assignment on list by title and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_api/web/lists/getByTitle(\'test\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment on list by id and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment on list by url and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetList(\'%2Fsites%2Fdocuments\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listUrl: 'sites/documents',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment on list get principal id by upn', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getUserByEmail').resolves(userResponse);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'User cannot be found.';
    sinon.stub(spo, 'getUserByEmail').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    } as any), new CommandError(error));
  });

  it('add role assignment on list get principal id by group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getGroupByName').resolves(groupResponse);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when group does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'Group cannot be found';
    sinon.stub(spo, 'getGroupByName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    } as any), new CommandError(error));
  });

  it('add role assignment on list get role definition id by role definition name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getRoleDefinitionByName').resolves(roledefinitionResponse);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    });
  });

  it('correctly handles error when role definition does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'No roledefinition is found for Read';
    sinon.stub(spo, 'getRoleDefinitionByName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    } as any), new CommandError(error));
  });
});
