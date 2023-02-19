import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import spoGroupGetCommand from '../group/group-get.js';
import spoRoleDefinitionListCommand from '../roledefinition/roledefinition-list.js';
import spoUserGetCommand from '../user/user-get.js';
import command from './listitem-roleassignment-add.js';

describe(commands.LISTITEM_ROLEASSIGNMENT_ADD, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name,commands.LISTITEM_ROLEASSIGNMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', listItemId: 1, principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 'abc', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the principalId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the roleDefinitionId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11, roleDefinitionId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the roleDefinitionId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if listId and listTitle are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, listTitle: 'Documents', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId and listUrl are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listUrl: '/sites/Documents', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle and listUrl are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', listItemId: 1, listUrl: '/sites/Documents', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if principalId and upn are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', listItemId: 1, principalId: 11, upn: 'someaccount@tenant.onmicrosoft.com', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if principalId and groupName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', listItemId: 1, principalId: 11, groupName: 'someGroup', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if upn and groupName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', listItemId: 1, upn: 'someaccount@tenant.onmicrosoft.com', groupName: 'someGroup', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if roleDefinitionId and roleDefinitionName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', listItemId: 1, groupName: 'someGroup', roleDefinitionId: 1073741827, roleDefinitionName: 'readers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation neither roleDefinitionId nor roleDefinitionName is specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', listItemId: 1, groupName: 'someGroup' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation neither groupName nor principalId or upn is specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', listItemId: 1, roleDefinitionName: 'readers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation neither listTitle nor listId or listUrl is specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', groupName: 'someGroup', listItemId: 1, roleDefinitionName: 'readers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listItemId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', groupName: 'someGroup', listItemId: 'abc', roleDefinitionName: 'readers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listItemId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', groupName: 'someGroup', listItemId: 1, roleDefinitionName: 'readers' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('add role assignment to listitem in list by title and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === 'https://contoso.sharepoint.com/_api/web/lists/getByTitle(\'test\')/items(1)/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test',
        listItemId: 1,
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment to listitem in list by id and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === 'https://contoso.sharepoint.com/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return '';
      }

      throw 'Invalid request';
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 1,
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment to listitem in list by url and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === 'https://contoso.sharepoint.com/_api/web/GetList(\'%2Fsites%2Fdocuments\')/items(1)/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return '';
      }

      throw 'Invalid request';
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listUrl: 'sites/documents',
        listItemId: 1,
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment to listitem in list get principal id by upn', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === 'https://contoso.sharepoint.com/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoUserGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "i:0#.f|membership|someaccount@tenant.onmicrosoft.com","Title": "Some Account","PrincipalType": 1,"Email": "someaccount@tenant.onmicrosoft.com","Expiration": "","IsEmailAuthenticationGuestUser": false,"IsShareByEmailGuestUser": false,"IsSiteAdmin": true,"UserId": {"NameId": "1003200097d06dd6","NameIdIssuer": "urn:federation:microsoftonline"},"UserPrincipalName": "someaccount@tenant.onmicrosoft.com"}'
        };
      }

      throw new CommandError('Unknown case');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === 'https://contoso.sharepoint.com/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return '';
      }

      throw 'Invalid request';
    });

    const error = 'no user found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoUserGetCommand) {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment to listitem in list get principal id by group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === 'https://contoso.sharepoint.com/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        };
      }

      throw new CommandError('Unknown case');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when group does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === 'https://contoso.sharepoint.com/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return '';
      }

      throw 'Invalid request';
    });

    const error = 'no group found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupGetCommand) {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment to listitem in list get role definition id by role definition name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === 'https://contoso.sharepoint.com/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoRoleDefinitionListCommand) {
        return {
          stdout: '[{"BasePermissions": {"High": "2147483647","Low": "4294967295"},"Description": "Has full control.","Hidden": false,"Id": 1073741827,"Name": "Full Control","Order": 1,"RoleTypeKind": 5}]'
        };
      }

      throw new CommandError('Unknown case');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    });
  });

  it('correctly handles error when role definition does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === 'https://contoso.sharepoint.com/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return '';
      }

      throw 'Invalid request';
    });

    const error = 'no role definition found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoRoleDefinitionListCommand) {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    });
  });

  it('correctly handles error when adding role definition fails', async () => {
    const error = 'error in adding role definition';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url as string === 'https://contoso.sharepoint.com/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    } as any), new CommandError(error));
  });
});
