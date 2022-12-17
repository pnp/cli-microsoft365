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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
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
      request.post,
      Cli.prompt,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GROUP_MEMBER_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webURL is Invalid', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "InvalidWEBURL",
        groupId: groupId,
        userName: userName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
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
        debug: false,
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
        debug: false,
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
        debug: false,
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
        debug: false,
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
        debug: false,
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
        debug: false,
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
        debug: false,
        verbose: true,
        webUrl: webUrl,
        groupId: groupId,
        userName: userName,
        confirm: true
      }
    }), new CommandError('The user does not exist or is not unique.'));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});