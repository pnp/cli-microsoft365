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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./group-member-remove');

describe(commands.GROUP_MEMBER_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const UserRemovalJSONResponse =
  {
    "odata.null": true
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
      Cli.prompt
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

  it('Removes user from SharePoint group using Group ID - Without Confirmation Prompt (Debug)', async () => {
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
        userName: "Alex.Wilber@contoso.com",
        confirm: true
      }
    });
    assert(postStub.called);
  });

  it('Removes user from SharePoint group using Group Name - Without Confirmation Prompt (Debug)', async () => {
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
        userName: "Alex.Wilber@contoso.com",
        confirm: true
      }
    });
    assert(postStub.called);
  });

  it('Removes user from SharePoint group using Group ID - With Confirmation Prompt', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
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
        confirm: false
      }
    });
    assert(postStub.called);
  });

  it('Aborts Removal of user from SharePoint group using Group Id - With Confirmation Prompt', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));

    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        userName: "Alex.Wilber@contoso.com",
        confirm: false
      }
    });
    assert(postSpy.notCalled);
  });

  it('Correctly Handles Error when removing user from the group using Group Id', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.reject('The user does not exist or is not unique.');
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        userName: "Alex.Wilber@invalidcontoso.com",
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

  it('fails validation if neither groupId nor groupName is entered', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", userName: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupId is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "INVALIDGROUP", userName: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all the required options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 3, userName: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
