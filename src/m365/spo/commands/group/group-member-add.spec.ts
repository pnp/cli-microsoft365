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
const command: Command = require('./group-member-add');

describe(commands.GROUP_MEMBER_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const jsonSingleUser =
  {
    ErrorMessage: null,
    IconUrl: "https://contoso.sharepoint.com/sites/SiteA/_layouts/15/images/siteicon.png",
    InvitedUsers: null,
    Name: "Site A",
    PermissionsPageRelativeUrl: null,
    StatusCode: 0,
    UniquelyPermissionedUsers: [],
    Url: "https://contoso.sharepoint.com/sites/SiteA",
    UsersAddedToGroup: [
      {
        AllowedRoles: [
          0
        ],
        CurrentRole: 0,
        DisplayName: "Alex Wilber",
        Email: "Alex.Wilber@contoso.com",
        InvitationLink: null,
        IsUserKnown: true,
        Message: null,
        Status: true,
        User: "i:0#.f|membership|Alex.Wilber@contoso.com"
      }
    ]
  };

  const jsonGroupNotFound =
  {
    status: 404,
    statusText: "Not Found",
    error: {
      "odata.error": {
        code: "-2146232832, Microsoft.SharePoint.SPException",
        message: {
          lang: "en-US",
          value: "Group cannot be found."
        }
      }
    }
  };

  const jsonErrorResponseInvalidUsers =
  {
    ErrorMessage: "Couldn't resolve the users.",
    IconUrl: null,
    InvitedUsers: null,
    Name: null,
    PermissionsPageRelativeUrl: null,
    StatusCode: -9,
    UniquelyPermissionedUsers: null,
    Url: null,
    UsersAddedToGroup: null
  };

  const jsonGenericError =
  {
    ErrorMessage: "The selected permission level is not valid.",
    IconUrl: null,
    InvitedUsers: null,
    Name: null,
    PermissionsPageRelativeUrl: null,
    StatusCode: -63,
    UniquelyPermissionedUsers: null,
    Url: null,
    UsersAddedToGroup: null
  };

  const userInformation: any =
  {
    businessPhones: [],
    displayName: "Alex Wilber",
    givenName: "Alex Wilber",
    id: "59b75414-4511-4c65-86a3-b6f5cd806748",
    jobTitle: "",
    mail: "Alex.Wilber@contoso.com",
    mobilePhone: null,
    officeLocation: null,
    preferredLanguage: null,
    surname: "User",
    userPrincipalName: "Alex.Wilber@contoso.com"
  };

  const spoUserInformation: any =
  {
    Id: 9,
    IsHiddenInUI: false,
    LoginName: "i:0#.f|membership|Alex.Wilber@contoso.com",
    Title: "Alex Wilber",
    PrincipalType: 1,
    Email: "",
    Expiration: "",
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: false,
    UserId: {
      NameId: "10032002529a911c",
      NameIdIssuer: "urn:federation:microsoftonline"
    },
    UserPrincipalName: "Alex.Wilber@contoso.com"
  };

  const azureGroupInformation: any = {
    id: "56ca9023-3449-4e98-a96a-69e81a6f4983",
    deletedDateTime: null,
    classification: null,
    createdDateTime: "2022-12-08T09:17:53Z",
    creationOptions: [
      "Team",
      "ExchangeProvisioningFlags:3552"
    ],
    description: "azuregroupname",
    displayName: "azuregroupname",
    expirationDateTime: null,
    groupTypes: [
      "Unified"
    ],
    isAssignableToRole: null,
    mail: "azuregroupname@ordidev.onmicrosoft.com",
    mailEnabled: true,
    mailNickname: "azuregroupname",
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
      "SPO:SPO_0cfa47ca-d4f4-44d6-ac7d-b83e815b3184@SPO_0cac6cda-2e04-4a3d-9c16-9c91470d7022",
      "SMTP:chipolata@ordidev.onmicrosoft.com"
    ],
    renewedDateTime: "2022-12-08T09:17:53Z",
    resourceBehaviorOptions: [
      "HideGroupInOutlook",
      "SubscribeMembersToCalendarEventsDisabled",
      "WelcomeEmailDisabled"
    ],
    resourceProvisioningOptions: [
      "Team"
    ],
    securityEnabled: false,
    securityIdentifier: "S-1-12-1-1456115747-1318597705-3899222697-2202627866",
    theme: null,
    visibility: "Public",
    onPremisesProvisioningErrors: []
  };

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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Cli.executeCommandWithOutput
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
    assert.strictEqual(command.name.startsWith(commands.GROUP_MEMBER_ADD), true);
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
        userName: "Alex.Wilber@contoso.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both groupId and groupName options are not passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        userName: "Alex.Wilber@contoso.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userName and email options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        email: "Alex.Wilber@contoso.com",
        userName: "Alex.Wilber@contoso.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userName and userId options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userId: 5,
        userName: "Alex.Wilber@contoso.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userName and aadGroupId options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        email: "Alex.Wilber@contoso.com",
        aadGroupId: "56ca9023-3449-4e98-a96a-69e81a6f4983"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userName and aadGroupName options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userId: 5,
        aadGroupName: "Azure AD Group name"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userId and email options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userId: 5,
        email: "Alex.Wilber@contoso.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName, email, userId, aadGroupId or aadGroupName options are not passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webURL is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "InvalidWEBURL", groupId: 32, userName: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupID is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "NOGROUP", userName: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, userId: "9,invalidUserId" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, userName: "Alex.Wilber@contoso.com,9" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if email is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, email: "Alex.Wilber@contoso.com,9" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if aadGroupId is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, aadGroupId: "56ca9023-3449-4e98-a96a-69e81a6f4983,9" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all the required options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, userName: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['DisplayName', 'Email']);
  });

  it('adds user to a SharePoint Group by groupId and userName', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(userInformation),
      stderr: ''
    }));

    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')`) {
        return Promise.resolve({
          Id: 32
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userName: "Alex.Wilber@contoso.com"
      }
    });
    assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
  });

  it('adds user to a SharePoint Group by groupId and userId (Debug)', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(spoUserInformation),
      stderr: ''
    }));

    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')`) {
        return Promise.resolve({
          Id: 32
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userId: 9
      }
    });
    assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
  });

  it('adds user to a SharePoint Group by groupId and userName (Debug)', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(userInformation),
      stderr: ''
    }));

    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')`) {
        return Promise.resolve({
          Id: 32
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userName: "Alex.Wilber@contoso.com"
      }
    });
    assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
  });

  it('adds user to a SharePoint Group by groupName and email (DEBUG)', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(userInformation),
      stderr: ''
    }));

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetByName(`) > -1) {
        return Promise.resolve({
          Id: 7
        });
      }
      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Contoso Site Owners",
        email: "Alex.Wilber@contoso.com"
      }
    });
    assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
  });

  it('adds user to a SharePoint Group by groupId and aadGroupId (Debug)', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(azureGroupInformation),
      stderr: ''
    }));

    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')`) {
        return Promise.resolve({
          Id: 32
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        aadGroupId: "56ca9023-3449-4e98-a96a-69e81a6f4983"
      }
    });
    assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
  });

  it('adds user to a SharePoint Group by groupId and aadGroupName (Debug)', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(azureGroupInformation),
      stderr: ''
    }));

    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')`) {
        return Promise.resolve({
          Id: 32
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        aadGroupName: "Azure AD Group name"
      }
    });
    assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
  });

  it('fails to get group when does not exists', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetByName('`) > -1) {
        return Promise.resolve({});
      }
      return Promise.reject('The specified group not exist in the SharePoint site');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Contoso Site Owners",
        email: "Alex.Wilber@contoso.com"
      }
    }), new CommandError(`The specified group does not exist in the SharePoint site`));
  });

  it('handles error when adding user to a SharePoint Group - Invalid Group', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('`) > -1) {
        return Promise.reject(jsonGroupNotFound);
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);

    });
    await assert.rejects(command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 99999999,
        userName: "Alex.Wilber@contoso.com"
      }
    }), new CommandError("Group cannot be found."));
  });

  it('handles error when adding user to a SharePoint Group ID - Username Does Not exist', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('4')`) {
        return Promise.resolve({
          Id: 4
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.reject({
      error: `Resource 'Alex.Wilber@invalidcontoso.onmicrosoft.com' does not exist or one of its queried reference-property objects are not present.`,
      stderr: `Resource 'Alex.Wilber@invalidcontoso.onmicrosoft.com' does not exist or one of its queried reference-property objects are not present. stderr`
    }));

    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonErrorResponseInvalidUsers);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        userName: "Alex.Wilber@invalidcontoso.onmicrosoft.com"
      }
    }), new CommandError(`Users not added to the group because the following users don't exist: Alex.Wilber@invalidcontoso.onmicrosoft.com`));
  });

  it('Handles generic error when adding user to a SharePoint Group by groupId and userName', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')`) {
        return Promise.resolve({
          Id: 32
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(userInformation),
      stderr: ''
    }));

    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonGenericError);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userName: "Alex.Wilber@contoso.com"
      }
    }), new CommandError(`The selected permission level is not valid.`));
  });

  it('Handles generic error when adding user to a SharePoint Group by groupId and userId', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')`) {
        return Promise.resolve({
          Id: 32
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.reject({
      error: `User cannot be found`,
      stderr: `User cannot be found`
    }));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userId: 19
      }
    }), new CommandError(`Users not added to the group because the following users don't exist: 19`));
  });
});
