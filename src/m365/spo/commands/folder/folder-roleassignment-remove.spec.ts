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
import command, { options } from './folder-roleassignment-remove.js';
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

describe(commands.FOLDER_ROLEASSIGNMENT_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let requests: any[];
  let promptIssued: boolean = false;
  const userResponse = {
    Id: 11,
    IsHiddenInUI: false,
    LoginName: 'i:0#.f|membership|someaccount@tenant.onmicrosoft.com',
    Title: 'Some Account',
    PrincipalType: 1,
    Email: 'someaccount@tenant.onmicrosoft.com',
    Expiration: '',
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: true,
    UserId: {
      NameId: '1003200097d06dd6',
      NameIdIssuer: 'urn:federation:microsoftonline'
    },
    UserPrincipalName: 'someaccount@tenant.onmicrosoft.com'
  };
  const groupResponse = {
    Id: 11,
    IsHiddenInUI: false,
    LoginName: 'otherGroup',
    Title: 'otherGroup',
    PrincipalType: 8,
    AllowMembersEditMembership: false,
    AllowRequestToJoinLeave: false,
    AutoAcceptRequestToJoinLeave: false,
    Description: '',
    OnlyAllowMembersViewMembership: true,
    OwnerTitle: 'Some Account',
    RequestToJoinLeaveEmailSetting: null
  };
  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      spo.getUserByEmail,
      spo.getGroupByName,
      cli.getSettingWithDefaultValue,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue,
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
    assert.strictEqual(command.name, commands.FOLDER_ROLEASSIGNMENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', folderUrl: '/Shared Documents/FolderPermission', principalId: 11 });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11 });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the principalId option is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: -5 });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `Specified principalId '-5' is not a valid number.`);
  });

  it('passes validation if the principalId option is a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11 });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if principalId and upn are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, upn: 'someaccount@tenant.onmicrosoft.com' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if principalId and groupName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, groupName: 'someGroup' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if upn and groupName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', upn: 'someaccount@tenant.onmicrosoft.com', groupName: 'someGroup' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither upn nor principalId or groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if folderUrl is not specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', upn: 'someaccount@tenant.onmicrosoft.com' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the entraGroupId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', entraGroupId: 'invalid' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'invalid' is not a valid GUID for option entraGroupId.`);
  });

  it('passes validation if the entraGroupId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', entraGroupId: '37455d5c-e466-4e49-8eba-808b5acec21b' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('remove role assignment from folder by folderUrl', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        principalId: 11,
        force: true
      })
    });
  });

  it('remove role assignment from folder and get principal id by upn', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getUserByEmail').resolves(userResponse);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        upn: 'someaccount@tenant.onmicrosoft.com',
        force: true
      })
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'no user found';
    sinon.stub(spo, 'getUserByEmail').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        upn: 'someaccount@tenant.onmicrosoft.com',
        force: true
      })
    }), new CommandError(error));
  });

  it('remove role assignment from folder and get principal id by group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getGroupByName').resolves(groupResponse);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup',
        force: true
      })
    });
  });

  it('removes role assignment from folder and gets principal id by Entra group ID', async () => {
    sinon.stub(entraGroup, 'getGroupById').withArgs(graphGroup.id).resolves(graphGroup);
    sinon.stub(spo, 'ensureEntraGroup').withArgs('https://contoso.sharepoint.com', graphGroup).resolves(entraGroupResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        entraGroupId: '27ae47f1-48f1-46f3-980b-d3c1470e398d',
        force: true
      })
    });
  });

  it('removes role assignment from folder and gets principal id by Entra group name', async () => {
    sinon.stub(entraGroup, 'getGroupByDisplayName').withArgs(graphGroup.displayName).resolves(graphGroup);
    sinon.stub(spo, 'ensureEntraGroup').withArgs('https://contoso.sharepoint.com', graphGroup).resolves(entraGroupResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        entraGroupName: 'Marketing',
        force: true
      })
    });
  });

  it('correctly handles error when group does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'no group found';
    sinon.stub(spo, 'getGroupByName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup',
        force: true
      })
    }), new CommandError(error));
  });

  it('aborts removing role assignment when prompt not confirmed', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup'
      })
    });

    assert(requests.length === 0);
  });

  it('prompts before removing role assignment when confirmation argument not passed', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup'
      })
    });
    assert(promptIssued);
  });

  it('removes role assignment when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getGroupByName').resolves(groupResponse);

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup'
      })
    });
  });
});
