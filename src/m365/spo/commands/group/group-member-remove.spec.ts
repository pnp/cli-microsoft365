import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as SpoGroupMemberListCommand from './group-member-list';
const command: Command = require('./group-member-remove');

describe(commands.GROUP_MEMBER_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const webUrl = 'https://contoso.sharepoint.com/sites/SiteA';
  const groupId = 4;
  const groupName = 'Site A Visitors';
  const userName = 'Alex.Wilber@contoso.com';
  const email = 'Alex.Wilber@contoso.com';
  const userId = 14;

  const spoGroupMemberListCommandOutput = `[{ "Id": 13, "IsHiddenInUI": false, "LoginName": "c:0t.c|tenant|4b468129-3b44-4414-bd45-aa5bde29df2f", "Title": "Azure AD Security Group 2", "PrincipalType": 1, "Email": "", "Expiration": "", "IsEmailAuthenticationGuestUser": false, "IsShareByEmailGuestUser": false, "IsSiteAdmin": false, "UserId": null, "UserPrincipalName": null },{ "Id": 13, "IsHiddenInUI": false, "LoginName": "c:0t.c|tenant|3f10f4af-8704-4394-80c0-ee8cef5eae27", "Title": "Azure AD Security Group", "PrincipalType": 1, "Email": "", "Expiration": "", "IsEmailAuthenticationGuestUser": false, "IsShareByEmailGuestUser": false, "IsSiteAdmin": false, "UserId": null, "UserPrincipalName": null }, { "Id": 17, "IsHiddenInUI": false, "LoginName": "c:0o.c|federateddirectoryclaimprovider|5786b8e8-c495-4734-b345-756733960730", "Title": "Office 365 Group", "PrincipalType": 4, "Email": "office365group@contoso.onmicrosoft.com", "Expiration": "", "IsEmailAuthenticationGuestUser": false, "IsShareByEmailGuestUser": false, "IsSiteAdmin": false, "UserId": null, "UserPrincipalName": null }]`;
  const UserRemovalJSONResponse =
  {
    "odata.null": true
  };

  const userInformation: any =
  {
    businessPhones: [],
    displayName: "Alex Wilber",
    givenName: "Alex Wilber",
    id: "59b75414-4511-4c65-86a3-b6f5cd806748",
    jobTitle: "",
    mail: email,
    mobilePhone: null,
    officeLocation: null,
    preferredLanguage: null,
    surname: "User",
    userPrincipalName: email
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Cli.prompt,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GROUP_MEMBER_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('Removes Azure AD group from SharePoint group using Azure AD Group Name', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoGroupMemberListCommand) {
        return ({
          stdout: spoGroupMemberListCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    const postStub = sinon.stub(request, 'post').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetByName') > -1) {
        return Promise.resolve(UserRemovalJSONResponse);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
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
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoGroupMemberListCommand) {
        return ({
          stdout: spoGroupMemberListCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    const postStub = sinon.stub(request, 'post').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetByName') > -1) {
        return Promise.resolve(UserRemovalJSONResponse);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Site A Visitors",
        aadGroupId: "5786b8e8-c495-4734-b345-756733960730",
        confirm: true
      }
    });
    assert(postStub.called);
  });

  it('Removes Azure AD group from SharePoint group using Azure AD Group ID and SharePoint Group ID', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoGroupMemberListCommand) {
        return ({
          stdout: spoGroupMemberListCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    const postStub = sinon.stub(request, 'post').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.resolve(UserRemovalJSONResponse);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
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
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoGroupMemberListCommand) {
        return ({
          stdout: spoGroupMemberListCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Site A Visitors",
        aadGroupName: "Not existing group",
        confirm: true
      }
    }), new CommandError('The Azure AD group Not existing group is not found in SharePoint group Site A Visitors'));
  });

  it('Removes user from SharePoint group using Group ID - Without Confirmation Prompt', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.resolve(UserRemovalJSONResponse);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        userName: "Alex.Wilber@contoso.com",
        confirm: true
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
      const loginName: string = `i:0#.f|membership|${userName}`;
      if (opts.url === `${webUrl}/_api/web/sitegroups/GetById('${groupId}')/users/removeByLoginName(@LoginName)?@LoginName='${formatting.encodeQueryParameter(loginName)}'`) {
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
        confirm: true
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
      const loginName: string = `i:0#.f|membership|${userName}`;
      if (opts.url === `${webUrl}/_api/web/sitegroups/GetById('${groupId}')/users/removeByLoginName(@LoginName)?@LoginName='${formatting.encodeQueryParameter(loginName)}'`) {
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
        confirm: false
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
      if (opts.url === `${webUrl}/_api/web/sitegroups/GetById('${groupId}')/users/removeById(${userId})`) {
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
        confirm: false
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupId and email when confirm option not passed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(userInformation),
      stderr: ''
    }));

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      const loginName: string = `i:0#.f|membership|${userName}`;
      if (opts.url === `${webUrl}/_api/web/sitegroups/GetById('${groupId}')/users/removeByLoginName(@LoginName)?@LoginName='${formatting.encodeQueryParameter(loginName)}'`) {
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
        confirm: false
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupName and userId with confirm option', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/sitegroups/GetByName('${formatting.encodeQueryParameter(groupName)}')/users/removeById(${userId})`) {
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
        confirm: true
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupName and email with confirm option', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(userInformation),
      stderr: ''
    }));

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      const loginName: string = `i:0#.f|membership|${userName}`;
      if (opts.url === `${webUrl}/_api/web/sitegroups/GetByName('${formatting.encodeQueryParameter(groupName)}')/users/removeByLoginName(@LoginName)?@LoginName='${formatting.encodeQueryParameter(loginName)}'`) {
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
        confirm: true
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
        confirm: false
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
        confirm: true
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