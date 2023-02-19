import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as SpoFileGetCommand from './file-get';
import * as SpoRoleDefinitionListCommand from '../roledefinition/roledefinition-list';
import * as SpoUserGetCommand from '../user/user-get';
import * as SpoGroupGetCommand from '../group/group-get';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
const command: Command = require('./file-roleassignment-add');


describe(commands.FILE_ROLEASSIGNMENT_ADD, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const fileUrl = '/sites/project-x/documents/Test1.docx';
  const fileId = 'b2307a39-e878-458b-bc90-03bc578531d6';

  let log: any[];
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
      Cli.executeCommandWithOutput,
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
    assert.strictEqual(command.name.startsWith(commands.FILE_ROLEASSIGNMENT_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, groupName: 'Group name A', roleDefinitionName: 'Read' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'foo', groupName: 'Group name A', roleDefinitionName: 'Read' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a valid number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, principalId: 'NaN', roleDefinitionName: 'Read' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the roleDefinitionId option is not a valid number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, groupName: 'Group name A', roleDefinitionId: 'NaN' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if no roledefinition is passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, principalId: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and fileId are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, groupName: 'Group name A', roleDefinitionName: 'Read' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly handles error when adding file role assignment', async () => {
    const errorMessage = 'request rejected';
    sinon.stub(request, 'post').callsFake(async () => { throw errorMessage; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        principalId: 10,
        roleDefinitionId: 1073741827
      }
    }), new CommandError(errorMessage));
  });

  it('correctly adds role assignment specifying principalId and role definition name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2Fdocuments%2FTest1.docx')/ListItemAllFields/roleassignments/addroleassignment(principalid='10',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoRoleDefinitionListCommand) {
        return Promise.resolve({
          stdout: '[{"BasePermissions": {"High": "2147483647","Low": "4294967295"},"Description": "Has full control.","Hidden": false,"Id": 1073741827,"Name": "Full Control","Order": 1,"RoleTypeKind": 5}]'
        });
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        principalId: 10,
        roleDefinitionName: 'Full Control'
      }
    });
  });

  it('correctly adds role assignment specifying principalId and role definition name, retrieving file by the ID', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2Fdocuments%2FTest1.docx')/ListItemAllFields/roleassignments/addroleassignment(principalid='10',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoFileGetCommand) {
        return ({
          stdout: '{"LinkingUri": "https://contoso.sharepoint.com/sites/project-x/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866","Name": "Test1.docx","ServerRelativeUrl": "/sites/project-x/documents/Test1.docx","UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6"}'
        });
      }
      if (command === SpoRoleDefinitionListCommand) {
        return {
          stdout: '[{"BasePermissions": {"High": "2147483647","Low": "4294967295"},"Description": "Has full control.","Hidden": false,"Id": 1073741827,"Name": "Full Control","Order": 1,"RoleTypeKind": 5}]'
        };
      }

      throw 'Unknown case';
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileId: fileId,
        principalId: 10,
        roleDefinitionName: 'Full Control'
      }
    });
  });

  it('correctly adds role assignment specifying upn and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2Fdocuments%2FTest1.docx')/ListItemAllFields/roleassignments/addroleassignment(principalid='11',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoUserGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "i:0#.f|membership|someaccount@tenant.onmicrosoft.com","Title": "Some Account","PrincipalType": 1,"Email": "someaccount@tenant.onmicrosoft.com","Expiration": "","IsEmailAuthenticationGuestUser": false,"IsShareByEmailGuestUser": false,"IsSiteAdmin": true,"UserId": {"NameId": "1003200097d06dd6","NameIdIssuer": "urn:federation:microsoftonline"},"UserPrincipalName": "someaccount@tenant.onmicrosoft.com"}'
        };
      }

      throw 'Unknown case';
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    const error = 'no user found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoUserGetCommand) {
        throw error;
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    }), new CommandError('no user found'));
  });

  it('correctly adds role assignment specifying groupName and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2Fdocuments%2FTest1.docx')/ListItemAllFields/roleassignments/addroleassignment(principalid='5',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoGroupGetCommand) {
        return {
          stdout: '{"Id": 5,"IsHiddenInUI": false,"LoginName": "Group A","Title": "Group A","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        };
      }

      throw 'Unknown case';
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        groupName: 'Group A',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when role definition does not exist', async () => {
    const error = 'no role definition found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command) => {
      if (command === SpoRoleDefinitionListCommand) {
        throw error;
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        groupName: 'Group A',
        roleDefinitionName: 'Non-existing Role Definition'
      }
    }), new CommandError('no role definition found'));
  });

  it('correctly handles error when group does not exist', async () => {
    const error = 'no group found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoGroupGetCommand) {
        throw error;
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        groupName: 'Group A',
        roleDefinitionId: 1073741827
      }
    }), new CommandError('no group found'));
  });
});
