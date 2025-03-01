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
import command from './list-roleassignment-remove.js';
import { settingsNames } from '../../../../settingsNames.js';
import { spo } from '../../../../utils/spo.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

describe(commands.LIST_ROLEASSIGNMENT_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  const webUrl = 'https://contoso.sharepoint.com/sites/Marketing';

  const userResponse = {
    Id: 11,
    IsHiddenInUI: false,
    LoginName: "i:0#.f|membership|someaccount@tenant.onmicrosoft.com",
    Title: "Some Account",
    PrincipalType: 1,
    Email: "someaccount@tenant.onmicrosoft.com",
    Expiration: "",
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: true,
    UserId: {
      NameId: "1003200097d06dd6",
      NameIdIssuer: "urn:federation:microsoftonline"
    },
    UserPrincipalName: "someaccount@tenant.onmicrosoft.com"
  };

  const groupResponse = {
    Id: 12,
    IsHiddenInUI: false,
    LoginName: "otherGroup",
    Title: "otherGroup",
    PrincipalType: 8,
    AllowMembersEditMembership: false,
    AllowRequestToJoinLeave: false,
    AutoAcceptRequestToJoinLeave: false,
    Description: "",
    OnlyAllowMembersViewMembership: true,
    OwnerTitle: "Some Account",
    RequestToJoinLeaveEmailSetting: null
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
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => settingName === settingsNames.prompt ? false : defaultValue);
    sinon.stub(entraGroup, 'getGroupById').withArgs(graphGroup.id).resolves(graphGroup);
    sinon.stub(entraGroup, 'getGroupByDisplayName').withArgs(graphGroup.displayName).resolves(graphGroup);
    sinon.stub(spo, 'ensureEntraGroup').withArgs(webUrl, graphGroup).resolves(entraGroupResponse);
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      spo.ensureUser,
      spo.getGroupByName,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_ROLEASSIGNMENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', principalId: 11 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the principalId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the upn option is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', upn: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the upn option is a valid user principal name', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', upn: 'john.doe@contoso.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the entraGroupId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', entraGroupId: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the entraGroupId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', entraGroupId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('removes role assignment from list by title', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('test')/roleassignments/removeroleassignment(principalid='11')`) {
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
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('removes role assignment from list by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/removeroleassignment(principalid='11')`) {
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
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('removes role assignment from list by url', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetList(\'%2Fsites%2Fdocuments\')/roleassignments/removeroleassignment(principalid='11')`) {
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
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('removes role assignment from list get principal id by upn', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/removeroleassignment(principalid='11')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'ensureUser').withArgs('https://contoso.sharepoint.com', 'someaccount@tenant.onmicrosoft.com').resolves(userResponse);

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        upn: 'someaccount@tenant.onmicrosoft.com',
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('correctly handles error when upn does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/removeroleassignment(principalid='11')`) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'User was not found';
    sinon.stub(spo, 'ensureUser').rejects({ error: { 'odata.error': { message: { value: error } } } });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        upn: 'someaccount@tenant.onmicrosoft.com',
        force: true
      }
    } as any), new CommandError(error));
  });

  it('remove role assignment from list get principal id by group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/removeroleassignment(principalid='12')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getGroupByName').withArgs('https://contoso.sharepoint.com', 'someGroup', logger, true).resolves(groupResponse);

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup',
        force: true
      }
    });
  });

  it('correctly handles error when group does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/removeroleassignment(principalid='12')`) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'no group found';
    sinon.stub(spo, 'getGroupByName').rejects({ error: { 'odata.error': { message: { value: error } } } });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup',
        force: true
      }
    }), new CommandError(error));
  });

  it('aborts removing role assignment when prompt not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup'
      }
    });

    assert(postStub.notCalled);
  });

  it('prompts before removing role assignment when confirmation argument not passed (Id)', async () => {
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listUrl: '/Lists/Tasks',
        groupName: 'someGroup'
      }
    });

    assert(promptIssued);
  });

  it('prompts before removing role assignment when confirmation argument not passed (Title)', async () => {
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'someList',
        groupName: 'someGroup'
      }
    });

    assert(promptIssued);
  });

  it('removes role assignment when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/removeroleassignment(principalid='12')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getGroupByName').withArgs('https://contoso.sharepoint.com', 'someGroup', logger, true).resolves(groupResponse);

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup'
      }
    });
  });

  it('removes role assignments for a Microsoft Entra group by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/removeroleassignment(principalid='15')`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        entraGroupId: graphGroup.id,
        force: true
      }
    });

    assert(postStub.calledOnce);
  });

  it('removes role assignments for a Microsoft Entra group by display name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists(guid'0CD891EF-AFCE-4E55-B836-FCE03286CCCF')/roleassignments/removeroleassignment(principalid='15')`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        entraGroupName: graphGroup.displayName,
        force: true
      }
    });

    assert(postStub.calledOnce);
  });
});
