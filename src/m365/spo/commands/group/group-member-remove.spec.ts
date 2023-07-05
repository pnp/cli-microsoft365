import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './group-member-remove.js';
import { spo } from '../../../../utils/spo.js';
import { aadUser } from '../../../../utils/aadUser.js';

describe(commands.GROUP_MEMBER_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const webUrl = 'https://contoso.sharepoint.com/sites/SiteA';
  const groupId = 4;
  const groupName = 'Site A Visitors';
  const userName = 'Alex.Wilber@contoso.com';
  const email = 'Alex.Wilber@contoso.com';
  const userId = 14;

  const groupMemberResponse = [
    {
      Id: 13,
      IsHiddenInUI: false,
      LoginName: 'c:0t.c|tenant|4b468129-3b44-4414-bd45-aa5bde29df2f',
      Title: 'Azure AD Security Group 2',
      PrincipalType: 1,
      Email: '',
      Expiration: '',
      IsEmailAuthenticationGuestUser: false,
      IsShareByEmailGuestUser: false,
      IsSiteAdmin: false,
      UserId: null,
      UserPrincipalName: null
    },
    {
      Id: 13,
      IsHiddenInUI: false,
      LoginName: 'c:0t.c|tenant|3f10f4af-8704-4394-80c0-ee8cef5eae27',
      Title: 'Azure AD Security Group',
      PrincipalType: 1,
      Email: '',
      Expiration: '',
      IsEmailAuthenticationGuestUser: false,
      IsShareByEmailGuestUser: false,
      IsSiteAdmin: false,
      UserId: null,
      UserPrincipalName: null
    }, {
      Id: 17,
      IsHiddenInUI: false,
      LoginName: 'c:0o.c|federateddirectoryclaimprovider|5786b8e8-c495-4734-b345-756733960730',
      Title: 'Office 365 Group',
      PrincipalType: 4,
      Email: 'office365group@contoso.onmicrosoft.com',
      Expiration: '',
      IsEmailAuthenticationGuestUser: false,
      IsShareByEmailGuestUser: false,
      IsSiteAdmin: false,
      UserId: null,
      UserPrincipalName: null
    }];
  const UserRemovalJSONResponse =
  {
    "odata.null": true
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      spo.getGroupMembersByGroupId,
      spo.getGroupMembersByGroupName,
      aadUser.getUpnByUserEmail,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_MEMBER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('Removes Azure AD group from SharePoint group using Azure AD Group Name', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    sinon.stub(spo, 'getGroupMembersByGroupName').resolves(groupMemberResponse);

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetByName') > -1) {
        return UserRemovalJSONResponse;
      }

      throw `Invalid request ${JSON.stringify(opts)
      }`;
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Site A Visitors",
        aadGroupName: "Azure AD Security Group"
      }
    });
    assert(postStub.called);
  });

  it('Removes Azure AD group from SharePoint group using Azure AD Group ID - Without Confirmation Prompt', async () => {
    sinon.stub(spo, 'getGroupMembersByGroupName').resolves(groupMemberResponse);

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetByName') > -1) {
        return UserRemovalJSONResponse;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Site A Visitors",
        aadGroupId: "5786b8e8-c495-4734-b345-756733960730",
        force: true
      }
    });
    assert(postStub.called);
  });

  it('Removes Azure AD group from SharePoint group using Azure AD Group ID and SharePoint Group ID', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    sinon.stub(spo, 'getGroupMembersByGroupId').resolves(groupMemberResponse);

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return UserRemovalJSONResponse;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        aadGroupId: "4b468129-3b44-4414-bd45-aa5bde29df2f"
      }
    });
    assert(postStub.called);
  });

  it('Throws error when Azure AD group not found', async () => {
    sinon.stub(spo, 'getGroupMembersByGroupName').resolves(groupMemberResponse);

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Site A Visitors",
        aadGroupName: "Not existing group",
        force: true
      }
    }), new CommandError('The Azure AD group Not existing group is not found in SharePoint group Site A Visitors'));
  });

  it('Removes user from SharePoint group using Group ID - Without Confirmation Prompt', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return UserRemovalJSONResponse;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        userName: "Alex.Wilber@contoso.com",
        force: true
      }
    });
    assert(postStub.called);
  });

  it('fails validation if the userName is not a valid UPN.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: groupId,
        userName: "no-an-email"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the email is not valid.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: groupId,
        email: "no-an-email"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupId and groupName is entered', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: groupId,
        groupName: groupName,
        userName: userName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither groupId nor groupName is entered', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        userName: userName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userId and email options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: groupId,
        userId: userId,
        email: email
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userName and email options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: groupId,
        email: email,
        userName: userName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userName and userId options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: groupId,
        userId: userId,
        userName: userName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both email and userId options are passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: groupId,
        email: email,
        userId: userId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName, email, and userId options are not passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: groupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupId is Invalid', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: "INVALIDGROUP",
        userName: userName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is Invalid', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: groupId,
        userId: "INVALIDUSER"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all the required options are specified', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        groupId: groupId,
        userName: userName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('removes user from SharePoint group by groupId and userName with confirm option (debug)', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      const loginName: string = `i: 0#.f | membership | ${userName}`;
      if (opts.url === `${webUrl} / _api / web / sitegroups / GetById('${groupId}') / users / removeByLoginName(@LoginName) ? @LoginName = '${formatting.encodeQueryParameter(loginName)}'`) {
        return UserRemovalJSONResponse;
      }

      return `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        webUrl: webUrl,
        groupId: groupId,
        userName: userName,
        force: true
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupId and userName when confirm option not passed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      const loginName: string = `i: 0#.f | membership | ${userName}`;
      if (opts.url === `${webUrl} / _api / web / sitegroups / GetById('${groupId}') / users / removeByLoginName(@LoginName) ? @LoginName = '${formatting.encodeQueryParameter(loginName)}'`) {
        return UserRemovalJSONResponse;
      }

      return `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        groupId: groupId,
        userName: userName,
        force: false
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupId and userId when confirm option not passed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl} / _api / web / sitegroups / GetById('${groupId}') / users / removeById(${userId})`) {
        return UserRemovalJSONResponse;
      }

      return `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        groupId: groupId,
        userId: userId,
        force: false
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupId and email when confirm option not passed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(aadUser, 'getUpnByUserEmail').resolves(email);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      const loginName: string = `i: 0#.f | membership | ${userName}`;
      if (opts.url === `${webUrl} / _api / web / sitegroups / GetById('${groupId}') / users / removeByLoginName(@LoginName) ? @LoginName = '${formatting.encodeQueryParameter(loginName)}'`) {
        return UserRemovalJSONResponse;
      }

      return `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        groupId: groupId,
        email: email,
        force: false
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupName and userId with confirm option', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl} / _api / web / sitegroups / GetByName('${formatting.encodeQueryParameter(groupName)}') / users / removeById(${userId})`) {
        return UserRemovalJSONResponse;
      }

      return `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        groupName: groupName,
        userId: userId,
        force: true
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupName and email with confirm option', async () => {
    sinon.stub(aadUser, 'getUpnByUserEmail').resolves(email);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      const loginName: string = `i: 0#.f | membership | ${userName}`;
      if (opts.url === `${webUrl} / _api / web / sitegroups / GetByName('${formatting.encodeQueryParameter(groupName)}') / users / removeByLoginName(@LoginName) ? @LoginName = '${formatting.encodeQueryParameter(loginName)}'`) {
        return UserRemovalJSONResponse;
      }

      return `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        groupName: groupName,
        email: email,
        force: true
      }
    });

    assert(postStub.called);
  });

  it('aborts removing user from SharePoint group when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        groupId: groupId,
        userName: userName,
        force: false
      }
    });
    assert(postSpy.notCalled);
  });

  it('correctly handles error when removing user from the group using groupId and userName', async () => {
    const loginName: string = `i:0#.f|membership|${userName}`;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/sitegroups/GetById('${groupId}')/users/removeByLoginName(@LoginName)?@LoginName='${formatting.encodeQueryParameter(loginName)}'`) {
        return Promise.reject('The user does not exist or is not unique.');
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        groupId: groupId,
        userName: userName,
        force: true
      }
    }), new CommandError('The user does not exist or is not unique.'));
  });

  it('fails validation if webURL is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "InvalidWEBURL", groupId: 4, userName: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupid and groupName is entered', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "4", groupName: "Site A Visitors", userName: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if aadGroupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3,
        aadGroupId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});