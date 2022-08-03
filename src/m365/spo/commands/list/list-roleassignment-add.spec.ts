import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
import * as SpoUserGetCommand from '../user/user-get';
import * as SpoGroupGetCommand from '../group/group-get';
import * as SpoRoleDefinitionListCommand from '../roledefinition/roledefinition-list';
const command: Command = require('./list-roleassignment-add');

describe(commands.LIST_ROLEASSIGNMENT_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_ROLEASSIGNMENT_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
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

  it('fails validation if listId and listTitle are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId and listUrl are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listUrl: '/sites/Documents', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle and listUrl are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', listUrl: '/sites/Documents', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if principalId and upn are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', principalId: 11, upn: 'someaccount@tenant.onmicrosoft.com', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if principalId and groupName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', principalId: 11, groupName: 'someGroup', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if upn and groupName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', upn: 'someaccount@tenant.onmicrosoft.com', groupName: 'someGroup', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if roleDefinitionId and roleDefinitionName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', groupName: 'someGroup', roleDefinitionId: 1073741827, roleDefinitionName: 'readers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('add role assignment on list by title and role definition id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/lists/getByTitle(\'test\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add role assignment on list by id and role definition id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add role assignment on list by url and role definition id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetList(\'%2Fsites%2Fdocuments\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listUrl: 'sites/documents',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add role assignment on list get principal id by upn', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoUserGetCommand) {
        return Promise.resolve({
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "i:0#.f|membership|someaccount@tenant.onmicrosoft.com","Title": "Some Account","PrincipalType": 1,"Email": "someaccount@tenant.onmicrosoft.com","Expiration": "","IsEmailAuthenticationGuestUser": false,"IsShareByEmailGuestUser": false,"IsSiteAdmin": true,"UserId": {"NameId": "1003200097d06dd6","NameIdIssuer": "urn:federation:microsoftonline"},"UserPrincipalName": "someaccount@tenant.onmicrosoft.com"}'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when upn does not exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    const error = 'no user found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoUserGetCommand) {
        return Promise.reject(error);
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(error)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add role assignment on list get principal id by group name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoGroupGetCommand) {
        return Promise.resolve({
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when group does not exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    const error = 'no group found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoGroupGetCommand) {
        return Promise.reject(error);
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(error)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add role assignment on list get role definition id by role definition name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoRoleDefinitionListCommand) {
        return Promise.resolve({
          stdout: '[{"BasePermissions": {"High": "2147483647","Low": "4294967295"},"Description": "Has full control.","Hidden": false,"Id": 1073741827,"Name": "Full Control","Order": 1,"RoleTypeKind": 5}]'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when role definition does not exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    const error = 'no role definition found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoRoleDefinitionListCommand) {
        return Promise.reject(error);
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(error)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});