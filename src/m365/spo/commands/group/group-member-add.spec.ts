import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './group-member-add.js';
import { settingsNames } from '../../../../settingsNames.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { Group } from '@microsoft/microsoft-graph-types';
import { CommandError } from '../../../../Command.js';

describe(commands.GROUP_MEMBER_ADD, () => {
  //#region API responses
  const userResponses =
    [
      {
        Id: 11,
        IsHiddenInUI: false,
        LoginName: 'i:0#.f|membership|adelev@contoso.onmicrosoft.com',
        Title: 'Adele Vance',
        PrincipalType: 1,
        Email: 'Adele.Vance@contoso.onmicrosoft.com',
        Expiration: '',
        IsEmailAuthenticationGuestUser: false,
        IsShareByEmailGuestUser: false,
        IsSiteAdmin: false,
        UserId: {
          NameId: '10032001f5ac2029',
          NameIdIssuer: 'urn:federation:microsoftonline'
        },
        UserPrincipalName: 'adelev@contoso.onmicrosoft.com'
      },
      {
        Id: 12,
        IsHiddenInUI: false,
        LoginName: 'i:0#.f|membership|johnd@contoso.onmicrosoft.com',
        Title: 'John Doe',
        PrincipalType: 1,
        Email: 'John.Doe@contoso.onmicrosoft.com',
        Expiration: '',
        IsEmailAuthenticationGuestUser: false,
        IsShareByEmailGuestUser: false,
        IsSiteAdmin: false,
        UserId: {
          NameId: '10032001f5ac2029',
          NameIdIssuer: 'urn:federation:microsoftonline'
        },
        UserPrincipalName: 'johnd@contoso.onmicrosoft.com'
      }
    ];

  const groupResponses =
    [
      {
        Id: 13,
        IsHiddenInUI: false,
        LoginName: 'c:0o.c|federateddirectoryclaimprovider|27ae47f1-48f1-46f3-980b-d3c2460e398d',
        Title: 'Marketing Members',
        PrincipalType: 4,
        Email: 'Marketing@contoso.onmicrosoft.com',
        Expiration: '',
        IsEmailAuthenticationGuestUser: false,
        IsShareByEmailGuestUser: false,
        IsSiteAdmin: false,
        UserId: null,
        UserPrincipalName: null
      },
      {
        Id: 14,
        IsHiddenInUI: false,
        LoginName: 'c:0o.c|federateddirectoryclaimprovider|27ae47f1-48f1-46f3-980b-d3c2460e398d',
        Title: 'HR Members',
        PrincipalType: 4,
        Email: 'HR@contoso.onmicrosoft.com',
        Expiration: '',
        IsEmailAuthenticationGuestUser: false,
        IsShareByEmailGuestUser: false,
        IsSiteAdmin: false,
        UserId: null,
        UserPrincipalName: null
      }
    ];

  const entraGroupResponses: Group[] = [
    {
      id: '75c621f8-8a1c-435a-827f-9b4e7917681b',
      deletedDateTime: null,
      classification: null,
      createdDateTime: '2023-07-20T14:43:57Z',
      description: 'Marketing',
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
        'SMTP:Marketing@contoso.onmicrosoft.com',
        'SPO:SPO_53dec431-9d4f-415b-b12b-010259d5b4e1@SPO_18c58817-3bc9-489d-ac63-f7264fb357e5'
      ],
      renewedDateTime: '2023-07-20T14:43:57Z',
      securityEnabled: false,
      securityIdentifier: 'S-1-12-1-1975919096-1130007068-1318813570-459806585',
      theme: null,
      visibility: 'Public',
      onPremisesProvisioningErrors: [],
      serviceProvisioningErrors: []
    },
    {
      id: 'e08e899f-ba40-4e91-ab36-44d4fbaa454e',
      deletedDateTime: null,
      classification: null,
      createdDateTime: '2024-01-05T22:29:40Z',
      description: null,
      displayName: 'IT Administrators',
      expirationDateTime: null,
      groupTypes: [],
      isAssignableToRole: null,
      mail: null,
      mailEnabled: false,
      mailNickname: 'd274ca34-f',
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
      proxyAddresses: [],
      renewedDateTime: '2024-01-05T22:29:40Z',
      securityEnabled: true,
      securityIdentifier: 'S-1-12-1-3767437727-1318173248-3561240235-1313188603',
      theme: null,
      visibility: null,
      onPremisesProvisioningErrors: [],
      serviceProvisioningErrors: []
    }
  ];
  //#endregion

  //#region Option values
  const webUrl = 'https://contoso.sharepoint.com/sites/Marketing';
  const spGroupName = 'Marketing Site Owners';
  const spGroupId = 3;
  const spUserIds = userResponses.map(u => u.Id);
  const userNames = userResponses.map(u => u.UserPrincipalName);
  const emails = userResponses.map(u => u.Email);
  const entraGroupIds = entraGroupResponses.map(g => g.id);
  const entraGroupNames = entraGroupResponses.map(g => g.displayName);
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => settingName === settingsNames.prompt ? false : defaultValue);
    sinon.stub(entraGroup, 'getGroupById').callsFake(async (id: string) => entraGroupResponses.find(g => g.id === id)!);
    sinon.stub(entraGroup, 'getGroupByDisplayName').callsFake(async (name: string) => entraGroupResponses.find(g => g.displayName === name)!);
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_MEMBER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both groupId and groupName options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        groupName: "Contoso Site Owners",
        userNames: "Alex.Wilber@contoso.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both groupId and groupName options are not passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        userNames: "Alex.Wilber@contoso.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userNames and emails options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        emails: "Alex.Wilber@contoso.com",
        userNames: "Alex.Wilber@contoso.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userNames and userIds options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userIds: 5,
        userNames: "Alex.Wilber@contoso.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both emails and entraGroupIds options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        emails: "Alex.Wilber@contoso.com",
        entraGroupIds: "56ca9023-3449-4e98-a96a-69e81a6f4983"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both emails and aadGroupIds options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        emails: "Alex.Wilber@contoso.com",
        aadGroupIds: "56ca9023-3449-4e98-a96a-69e81a6f4983"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userIds and entraGroupNames options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userIds: 5,
        entraGroupNames: "Microsoft Entra Group name"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userIds and aadGroupNames options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userIds: 5,
        aadGroupNames: "Microsoft Entra Group name"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userIds and emails options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userIds: 5,
        emails: "Alex.Wilber@contoso.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userNames, emails, userIds, entraGroupIds, aadGroupIds, entraGroupNames, or aadGroupNames options are not passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webURL is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "InvalidWEBURL", groupId: 32, userNames: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupID is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "NOGROUP", userNames: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userIds is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, userIds: "9,invalidUserId" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userNames is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, userNames: "Alex.Wilber@contoso.com,9" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if emails is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, emails: "Alex.Wilber@contoso.com,9" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if entraGroupIds is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, entraGroupIds: "56ca9023-3449-4e98-a96a-69e81a6f4983,9" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if aadGroupIds is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, aadGroupIds: "56ca9023-3449-4e98-a96a-69e81a6f4983,9" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all the required options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, userNames: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'UserPrincipalName']);
  });

  it('correctly logs deprecation warning for aadGroupIds option', async () => {
    const chalk = (await import('chalk')).default;
    const loggerErrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(request, 'post').resolves();

    await command.action(logger, { options: { webUrl: webUrl, groupName: spGroupName, aadGroupIds: entraGroupIds[0] } });
    assert.deepStrictEqual(loggerErrSpy.firstCall.firstArg, chalk.yellow(`Option 'aadGroupIds' is deprecated. Please use 'entraGroupIds' instead.`));

    sinonUtil.restore(loggerErrSpy);
  });

  it('correctly logs deprecation warning for aadGroupNames option', async () => {
    const chalk = (await import('chalk')).default;
    const loggerErrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(request, 'post').resolves();

    await command.action(logger, { options: { webUrl: webUrl, groupName: spGroupName, aadGroupNames: entraGroupNames[0] } });
    assert.deepStrictEqual(loggerErrSpy.firstCall.firstArg, chalk.yellow(`Option 'aadGroupNames' is deprecated. Please use 'entraGroupNames' instead.`));

    sinonUtil.restore(loggerErrSpy);
  });

  it('correctly logs result when adding users by userNames', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/SiteGroups/GetById(${spGroupId})/users`) {
        return userResponses[postStub.callCount - 1];
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        groupId: spGroupId,
        userNames: userNames.join(',')
      }
    });

    assert(loggerLogSpy.calledOnceWithExactly(userResponses));
  });

  it('correctly adds users to group by UPNs', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/SiteGroups/GetById(${spGroupId})/users`) {
        return userResponses[postStub.callCount - 1];
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        groupId: spGroupId,
        userNames: userNames.join(',')
      }
    });

    assert.deepStrictEqual(postStub.firstCall.args[0].data, { LoginName: 'i:0#.f|membership|adelev@contoso.onmicrosoft.com' });
    assert.deepStrictEqual(postStub.secondCall.args[0].data, { LoginName: 'i:0#.f|membership|johnd@contoso.onmicrosoft.com' });
  });

  it('correctly adds users to group by emails', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/SiteGroups/GetById(${spGroupId})/users`) {
        return userResponses[postStub.callCount - 1];
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        groupId: spGroupId,
        emails: emails.join(',')
      }
    });

    assert.deepStrictEqual(postStub.firstCall.args[0].data, { LoginName: 'i:0#.f|membership|Adele.Vance@contoso.onmicrosoft.com' });
    assert.deepStrictEqual(postStub.secondCall.args[0].data, { LoginName: 'i:0#.f|membership|John.Doe@contoso.onmicrosoft.com' });
  });

  it('correctly adds users to group by userIds', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/SiteGroups/GetById(${spGroupId})/users`) {
        return userResponses[postStub.callCount - 1];
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      for (const user of userResponses) {
        if (opts.url === `${webUrl}/_api/web/SiteUsers/GetById(${user.Id})?$select=LoginName`) {
          return { LoginName: user.LoginName };
        }
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        groupId: spGroupId,
        userIds: spUserIds.join(','),
        verbose: true
      }
    });

    assert.deepStrictEqual(postStub.firstCall.args[0].data, { LoginName: 'i:0#.f|membership|adelev@contoso.onmicrosoft.com' });
    assert.deepStrictEqual(postStub.secondCall.args[0].data, { LoginName: 'i:0#.f|membership|johnd@contoso.onmicrosoft.com' });
  });

  it('correctly logs result when adding groups', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/SiteGroups/GetById(${spGroupId})/users`) {
        return groupResponses[postStub.callCount - 1];
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        groupId: spGroupId,
        entraGroupIds: entraGroupIds.join(',')
      }
    });

    assert(loggerLogSpy.calledOnceWithExactly(groupResponses));
  });

  it('correctly adds users to group by entraGroupIds', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/SiteGroups/GetById(${spGroupId})/users`) {
        return groupResponses[postStub.callCount - 1];
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        groupId: spGroupId,
        entraGroupIds: entraGroupIds.join(','),
        verbose: true
      }
    });

    assert.deepStrictEqual(postStub.firstCall.args[0].data, { LoginName: `c:0o.c|federateddirectoryclaimprovider|${entraGroupResponses[0].id}` });
    assert.deepStrictEqual(postStub.secondCall.args[0].data, { LoginName: `c:0t.c|tenant|${entraGroupResponses[1].id}` });
  });

  it('correctly adds users to group by entraGroupNames', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/SiteGroups/GetById(${spGroupId})/users`) {
        return groupResponses[postStub.callCount - 1];
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        groupId: spGroupId,
        entraGroupNames: entraGroupNames.join(','),
        verbose: true
      }
    });

    assert.deepStrictEqual(postStub.firstCall.args[0].data, { LoginName: `c:0o.c|federateddirectoryclaimprovider|${entraGroupResponses[0].id}` });
    assert.deepStrictEqual(postStub.secondCall.args[0].data, { LoginName: `c:0t.c|tenant|${entraGroupResponses[1].id}` });
  });

  it('correctly handles error when adding users to group', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-2130575276, Microsoft.SharePoint.SPException',
          message: {
            lang: 'en-US',
            value: 'The user does not exist or is not unique.'
          }
        }
      }
    };

    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, groupId: spGroupId, userNames: userNames.join(',') } }),
      new CommandError(error.error['odata.error'].message.value));
  });
});
