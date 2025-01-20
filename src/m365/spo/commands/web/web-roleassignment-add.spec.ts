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
import spoRoleDefinitionListCommand from '../roledefinition/roledefinition-list.js';
import spoUserGetCommand from '../user/user-get.js';
import command from './web-roleassignment-add.js';
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
  Id: 11,
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

describe(commands.WEB_ROLEASSIGNMENT_ADD, () => {
  let log: any[];
  let logger: Logger;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.executeCommandWithOutput,
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
    assert.strictEqual(command.name, commands.WEB_ROLEASSIGNMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', principalId: 'abc', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the principalId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the roleDefinitionId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', principalId: 11, roleDefinitionId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the roleDefinitionId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the entaGroupId is not a valid guid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', entraGroupId: 'invalid', roleDefinitionId: '1073741827' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the entaGroupId is a valid guid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', entraGroupId: '27ae47f1-48f1-46f3-980b-d3c1470e398d', roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('add role assignment on web by role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_api/web/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment on web get principal id by upn', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
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
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
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
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    } as any), new CommandError(error));
  });

  it('add role assignment on web get principal id by group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
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
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when group does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
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
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    } as any), new CommandError(error));
  });

  it('add role assignment on web get role definition id by role definition name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoRoleDefinitionListCommand) {
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
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    });
  });

  it('adds role assignment on web by role definition id and Entra group ID', async () => {
    sinon.stub(entraGroup, 'getGroupById').withArgs(graphGroup.id).resolves(graphGroup);
    sinon.stub(spo, 'ensureEntraGroup').withArgs('https://contoso.sharepoint.com', graphGroup).resolves(entraGroupResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_api/web/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        entraGroupId: '27ae47f1-48f1-46f3-980b-d3c1470e398d',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('adds role assignment on web by role definition id and Entra group name', async () => {
    sinon.stub(entraGroup, 'getGroupByDisplayName').withArgs(graphGroup.displayName).resolves(graphGroup);
    sinon.stub(spo, 'ensureEntraGroup').withArgs('https://contoso.sharepoint.com', graphGroup).resolves(entraGroupResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_api/web/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        entraGroupName: 'Marketing',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when role definition does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'no role definition found';
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoRoleDefinitionListCommand) {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        principalId: 11,
        roleDefinitionName: 'Full Control'
      }
    } as any), new CommandError(error));
  });
});
