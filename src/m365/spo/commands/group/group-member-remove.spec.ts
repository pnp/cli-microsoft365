import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
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
import spoGroupMemberListCommand from './group-member-list.js';
import command from './group-member-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

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

  const spoGroupMemberListCommandOutput = `[{ "Id": 13, "IsHiddenInUI": false, "LoginName": "c:0t.c|tenant|4b468129-3b44-4414-bd45-aa5bde29df2f", "Title": "Microsoft Entra Security Group 2", "PrincipalType": 1, "Email": "", "Expiration": "", "IsEmailAuthenticationGuestUser": false, "IsShareByEmailGuestUser": false, "IsSiteAdmin": false, "UserId": null, "UserPrincipalName": null },{ "Id": 13, "IsHiddenInUI": false, "LoginName": "c:0t.c|tenant|3f10f4af-8704-4394-80c0-ee8cef5eae27", "Title": "Microsoft Entra Security Group", "PrincipalType": 1, "Email": "", "Expiration": "", "IsEmailAuthenticationGuestUser": false, "IsShareByEmailGuestUser": false, "IsSiteAdmin": false, "UserId": null, "UserPrincipalName": null }, { "Id": 17, "IsHiddenInUI": false, "LoginName": "c:0o.c|federateddirectoryclaimprovider|5786b8e8-c495-4734-b345-756733960730", "Title": "Office 365 Group", "PrincipalType": 4, "Email": "office365group@contoso.onmicrosoft.com", "Expiration": "", "IsEmailAuthenticationGuestUser": false, "IsShareByEmailGuestUser": false, "IsSiteAdmin": false, "UserId": null, "UserPrincipalName": null }]`;
  const UserRemovalJSONResponse = {
    "odata.null": true
  };

  const userInformation: any = {
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
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      request.get,
      request.post,
      cli.promptForConfirmation,
      cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_MEMBER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('Removes Entra group from SharePoint group using Entra Group Name', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupMemberListCommand) {
        return ({
          stdout: spoGroupMemberListCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

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
        entraGroupName: "Microsoft Entra Security Group"
      }
    });

    assert(postStub.called);
  });

  it('Removes Microsoft Entra group from SharePoint group using Microsoft Entra Group Name', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupMemberListCommand) {
        return ({
          stdout: spoGroupMemberListCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

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
        entraGroupName: "Microsoft Entra Security Group"
      }
    });

    assert(postStub.called);
  });

  it('Removes Entra group from SharePoint group using Entra Group ID - Without Confirmation Prompt', async () => {
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupMemberListCommand) {
        return ({
          stdout: spoGroupMemberListCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

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
        entraGroupId: "5786b8e8-c495-4734-b345-756733960730",
        force: true
      }
    });
    assert(postStub.called);
  });

  it('Removes Microsoft Entra group from SharePoint group using Microsoft Entra Group ID - Without Confirmation Prompt', async () => {
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupMemberListCommand) {
        return ({
          stdout: spoGroupMemberListCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

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
        entraGroupId: "5786b8e8-c495-4734-b345-756733960730",
        force: true
      }
    });
    assert(postStub.called);
  });

  it('Removes Entra group from SharePoint group using Entra Group ID and SharePoint Group ID', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupMemberListCommand) {
        return ({
          stdout: spoGroupMemberListCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

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
        entraGroupId: "4b468129-3b44-4414-bd45-aa5bde29df2f"
      }
    });
    assert(postStub.called);
  });

  it('Throws error when Microsoft Entra group not found', async () => {
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupMemberListCommand) {
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
        entraGroupName: "Not existing group",
        force: true
      }
    }), new CommandError('The Microsoft Entra group Not existing group is not found in SharePoint group Site A Visitors'));
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        userName: userName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userId and email options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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

  it('removes user from SharePoint group by groupId and userName with force option (debug)', async () => {
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
        force: true
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupId and userName when force option not passed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

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
        force: false
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupId and userId when force option not passed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

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
        force: false
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupId and email when force option not passed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
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
        force: false
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupName and userId with force option', async () => {
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
        force: true
      }
    });

    assert(postStub.called);
  });

  it('removes user from SharePoint group by groupName and email with force option', async () => {
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
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
        force: true
      }
    });

    assert(postStub.called);
  });

  it('aborts removing user from SharePoint group when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "4", groupName: "Site A Visitors", userName: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if entraGroupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3,
        entraGroupId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});