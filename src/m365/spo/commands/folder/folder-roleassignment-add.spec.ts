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
import spoRoleDefinitionFolderCommand from '../roledefinition/roledefinition-list.js';
import spoUserGetCommand from '../user/user-get.js';
import spoFolderGetCommand from './folder-get.js';
import command from './folder-roleassignment-add.js';

describe(commands.FOLDER_ROLEASSIGNMENT_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_ROLEASSIGNMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 'abc', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the principalId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the roleDefinitionId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, roleDefinitionId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the roleDefinitionId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if principalId and upn are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, upn: 'someaccount@tenant.onmicrosoft.com', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if principalId and groupName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, groupName: 'someGroup', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if upn and groupName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', upn: 'someaccount@tenant.onmicrosoft.com', groupName: 'someGroup', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if roleDefinitionId and roleDefinitionName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', groupName: 'someGroup', roleDefinitionId: 1073741827, roleDefinitionName: 'readers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation neither roleDefinitionId nor roleDefinitionName is specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', groupName: 'someGroup' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation neither groupName nor principalId or upn is specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', roleDefinitionName: 'readers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('add the role assignment to the specified folder based on the upn and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/breakroleinheritance(true)') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFolderGetCommand) {
        return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/Shared Documents/FolderPermission", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" };
      }

      if (command === spoUserGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "i:0#.f|membership|someaccount@tenant.onmicrosoft.com","Title": "Some Account","PrincipalType": 1,"Email": "someaccount@tenant.onmicrosoft.com","Expiration": "","IsEmailAuthenticationGuestUser": false,"IsShareByEmailGuestUser": false,"IsSiteAdmin": true,"UserId": {"NameId": "1003200097d06dd6","NameIdIssuer": "urn:federation:microsoftonline"},"UserPrincipalName": "someaccount@tenant.onmicrosoft.com"}'
        };
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add the role assignment to the specified root folder based on the principal id and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetList(\'%2FShared%20Documents\')/breakroleinheritance(true)') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetList(\'%2FShared%20Documents\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFolderGetCommand) {
        return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/Shared Documents/FolderPermission", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" };
      }
      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents\')/ListItemAllFields/breakroleinheritance(true)') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'no user found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFolderGetCommand) {
        return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/Shared Documents/FolderPermission", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" };
      }

      if (command === spoUserGetCommand) {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    } as any), new CommandError(error));
  });

  it('add the role assignment to the specified folder based on the group name and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/breakroleinheritance(true)') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return;
      }

      throw 'Invalid request';
    });


    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFolderGetCommand) {
        return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/Shared Documents/FolderPermission", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" };
      }

      if (command === spoGroupGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        };
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when group does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/breakroleinheritance(true)') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'no group found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFolderGetCommand) {
        return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/Shared Documents/FolderPermission", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" };
      }

      if (command === spoGroupGetCommand) {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    } as any), new CommandError(error));
  });

  it('add the role assignment to the specified folder based on the principal id and role definition name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/breakroleinheritance(true)') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return;
      }

      throw 'Invalid request';
    });


    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFolderGetCommand) {
        return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/Shared Documents/FolderPermission", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" };
      }

      if (command === spoRoleDefinitionFolderCommand) {
        return {
          stdout: '[{"BasePermissions": {"High": "2147483647","Low": "4294967295"},"Description": "Has full control.","Hidden": false,"Id": 1073741827,"Name": "Full Control","Order": 1,"RoleTypeKind": 5}]'
        };
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    });
  });

  it('correctly handles error when role definition does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/breakroleinheritance(true)') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') {
        return;
      }

      throw 'Invalid request';
    });

    const error = "The specified role definition name 'Full Control1' does not exist.";
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFolderGetCommand) {
        return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/Shared Documents/FolderPermission", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" };
      }

      if (command === spoRoleDefinitionFolderCommand) {
        return {
          stdout: '[{"BasePermissions": {"High": "2147483647","Low": "4294967295"},"Description": "Has full control.","Hidden": false,"Id": 1073741827,"Name": "Full Control","Order": 1,"RoleTypeKind": 5}]'
        };
      }
      throw new CommandError(error);
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        principalId: 11,
        roleDefinitionName: 'Full Control1'
      }
    } as any), new CommandError(error));
  });
});
