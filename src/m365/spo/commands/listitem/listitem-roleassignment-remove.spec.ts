import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
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
import spoUserGetCommand from '../user/user-get.js';
import command from './listitem-roleassignment-remove.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';

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
  mail: 'Marketing@contoso.onmicrosoft.com',
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
    'SMTP:Marketing@contoso.onmicrosoft.com'
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

describe(commands.LISTITEM_ROLEASSIGNMENT_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptIssued: boolean = false;

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
    requests = [];
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.executeCommandWithOutput,
      cli.promptForConfirmation,
      entraGroup.getGroupById,
      entraGroup.getGroupByDisplayName,
      spo.ensureEntraGroup
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_ROLEASSIGNMENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', listItemId: 1, principalId: 11 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the listItemId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 'abc', principalId: 11 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listItemId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the principalId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the entraGroupId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, entraGroupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the entraGroupId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: 1, entraGroupId: '37455d5c-e466-4e49-8eba-808b5acec21b' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('remove role assignment from list item get list by title', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_api/web/lists/getByTitle(\'test\')/items(1)/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
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
        force: true
      }
    });
  });

  it('remove role assignment from list item get list by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        principalId: 11,
        force: true
      }
    });
  });

  it('remove role assignment from list item get list by url', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetList(\'%2Fsites%2Fdocuments\')/items(1)/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listUrl: 'sites/documents',
        listItemId: 1,
        principalId: 11,
        force: true
      }
    });
  });

  it('remove role assignment from list item get principal id by upn', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
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
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        upn: 'someaccount@tenant.onmicrosoft.com',
        force: true
      }
    });
  });

  it('removes role assignment from list item get list by Entra group ID', async () => {
    sinon.stub(entraGroup, 'getGroupById').withArgs(graphGroup.id).resolves(graphGroup);
    sinon.stub(spo, 'ensureEntraGroup').withArgs('https://contoso.sharepoint.com', graphGroup).resolves(entraGroupResponse);


    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/removeroleassignment(principalid=\'15\')') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        entraGroupId: '27ae47f1-48f1-46f3-980b-d3c1470e398d',
        force: true
      }
    });
  });

  it('removes role assignment from list item get list by Entra group name', async () => {
    sinon.stub(entraGroup, 'getGroupByDisplayName').withArgs(graphGroup.displayName).resolves(graphGroup);
    sinon.stub(spo, 'ensureEntraGroup').withArgs('https://contoso.sharepoint.com', graphGroup).resolves(entraGroupResponse);


    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/removeroleassignment(principalid=\'15\')') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        entraGroupName: 'Marketing',
        force: true
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    const error = 'no user found';
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoUserGetCommand) {
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
        upn: 'someaccount@tenant.onmicrosoft.com',
        force: true
      }
    } as any), new CommandError(error));
  });

  it('remove role assignment from list item get principal id by group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
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
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        groupName: 'someGroup',
        force: true
      }
    });
  });

  it('correctly handles error when group does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    const error = 'no group found';
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupGetCommand) {
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
        groupName: 'someGroup',
        force: true
      }
    } as any), new CommandError(error));
  });

  it('aborts removing role assignment when prompt not confirmed', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        groupName: 'someGroup'
      }
    });

    assert(requests.length === 0);
  });

  it('prompts before removing role assignment when confirmation argument not passed (Id)', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        groupName: 'someGroup'
      }
    });
    assert(promptIssued);
  });

  it('prompts before removing role assignment when confirmation argument not passed (Title)', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'someList',
        listItemId: 1,
        groupName: 'someGroup'
      }
    });

    assert(promptIssued);
  });

  it('removes role assignment when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/items(1)/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        };
      }

      throw new CommandError('Unknown case');
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        listItemId: 1,
        groupName: 'someGroup'
      }
    });
  });
});
