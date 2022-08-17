import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./navigation-node-list');

describe(commands.NAVIGATION_NODE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.NAVIGATION_NODE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets nodes from the top navigation', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1) {
        return Promise.resolve({ value: [{ "Id": 2003, "IsDocLib": true, "IsExternal": false, "IsVisible": true, "ListTemplateType": 0, "Title": "Node 1", "Url": "/sites/team-a/SitePages/page1.aspx" }, { "Id": 2004, "IsDocLib": true, "IsExternal": false, "IsVisible": true, "ListTemplateType": 0, "Title": "Node 2", "Url": "/sites/team-a/SitePages/page2.aspx" }] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{ "Id": 2003, "Title": "Node 1", "Url": "/sites/team-a/SitePages/page1.aspx" }, { "Id": 2004, "Title": "Node 2", "Url": "/sites/team-a/SitePages/page2.aspx" }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets nodes from the quick launch', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/quicklaunch`) > -1) {
        return Promise.resolve({ value: [{ "Id": 2003, "IsDocLib": true, "IsExternal": false, "IsVisible": true, "ListTemplateType": 0, "Title": "Node 1", "Url": "/sites/team-a/SitePages/page1.aspx" }, { "Id": 2004, "IsDocLib": true, "IsExternal": false, "IsVisible": true, "ListTemplateType": 0, "Title": "Node 2", "Url": "/sites/team-a/SitePages/page2.aspx" }] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{ "Id": 2003, "Title": "Node 1", "Url": "/sites/team-a/SitePages/page1.aspx" }, { "Id": 2004, "Title": "Node 2", "Url": "/sites/team-a/SitePages/page2.aspx" }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1) {
        return Promise.reject({ error: 'An error has occurred' });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error (string error)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', location: 'TopNavigationBar' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified location is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is QuickLaunch and all required properties are present', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});