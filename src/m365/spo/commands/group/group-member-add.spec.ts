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
  let cli: Cli;
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

  const groupResponse = {
    Id: 32
  };

  before(() => {
    cli = Cli.getInstance();
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => {
      if (settingName === "prompt") { return false; }
      else {
        return defaultValue;
      }
    }));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
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
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
        return groupResponse;
      }

      throw 'Invalid request';
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
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/siteusers/GetById('9')?$select=AadObjectId`) {
        return Promise.resolve({
          AadObjectId: {
            NameId: '6cc1797e-5463-45ec-bb1a-b93ec198bab6',
            NameIdIssuer: 'urn:federation:microsoftonline'
          }
        });
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
        return Promise.resolve(groupResponse);
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
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
        return groupResponse;
      }

      throw 'Invalid request';
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq 'Alex.Wilber%40contoso.com'&$select=id`) {
        return Promise.resolve({ value: [{ id: "2056d2f6-3257-4253-8cfc-b73393e414e5" }] });
      }

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
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
        return groupResponse;
      }

      throw 'Invalid request';
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
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonSingleUser);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Azure%20AD%20Group%20name'&$select=id`) {
        return Promise.resolve({
          value: [{
            id: 'Group name'
          }]
        });
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
        return Promise.resolve(groupResponse);
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
    const errorMessage = 'The specified group does not exist in the SharePoint site';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetByName('`) > -1) {
        return Promise.reject({ error: { 'odata.error': { message: { value: errorMessage } } } });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Contoso Site Owners",
        email: "Alex.Wilber@contoso.com"
      }
    }), new CommandError(errorMessage));
  });

  it('handles generic error when adding user to a SharePoint Group by groupId and userName', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/siteusers/GetById('9')`) {
        return Promise.reject(`User not found`);
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
        return Promise.resolve(groupResponse);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userId: 9
      }
    }), new CommandError(`Resource '9' does not exist.`));
  });

  it('handles error when adding user to SharePoint Group group', async () => {
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return Promise.resolve(jsonGenericError);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
        return groupResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userName: 'Alex.Wilber@contoso.com'
      }
    }), new CommandError('The selected permission level is not valid.'));
  });
});
