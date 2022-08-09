import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError} from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./group-user-remove');

describe(commands.GROUP_USER_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const UserRemovalJSONResponse =
  {
    "odata.null": true
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GROUP_USER_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('Removes user from SharePoint group using Group ID - Without Confirmation Prompt', (done) => {
    sinon.stub(request, 'post').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.resolve(UserRemovalJSONResponse);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    command.action(logger, {
      options: {
        debug: false,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        userName : "Alex.Wilber@contoso.com",
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Removes user from SharePoint group using Group ID - Without Confirmation Prompt (Debug)', (done) => {
    sinon.stub(request, 'post').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.resolve(UserRemovalJSONResponse);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        userName : "Alex.Wilber@contoso.com",
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Removes user from SharePoint group using Group Name - Without Confirmation Prompt (Debug)', (done) => {
    sinon.stub(request, 'post').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetByName') > -1) {
        return Promise.resolve(UserRemovalJSONResponse);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Site A Visitors",
        userName : "Alex.Wilber@contoso.com",
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Removes user from SharePoint group using Group ID - With Confirmation Prompt', (done) => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    sinon.stub(request, 'post').callsFake(opts => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.resolve(UserRemovalJSONResponse);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    command.action(logger, {
      options: {
        debug: false,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        userName : "Alex.Wilber@contoso.com",
        confirm: false
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Aborts Removal of user from SharePoint group using Group Id - With Confirmation Prompt', (done) => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });

    command.action(logger, {
      options: {
        debug: false,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        userName : "Alex.Wilber@contoso.com",
        confirm: false
      }
    }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Correctly Handles Error when removing user from the group using Group Id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.reject('The user does not exist or is not unique.');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 4,
        userName : "Alex.Wilber@invalidcontoso.com",
        confirm: true
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError('The user does not exist or is not unique.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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