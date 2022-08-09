import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./navigation-node-add');

describe(commands.NAVIGATION_NODE_ADD, () => {
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
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.NAVIGATION_NODE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['url']);
  });

  it('adds new navigation node to the top navigation', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          Title: 'About',
          Url: '/sites/team-a/sitepages/about.aspx',
          IsExternal: false
        })) {
        return Promise.resolve(
          {
            "Id": 2001,
            "IsDocLib": true,
            "IsExternal": false,
            "IsVisible": true,
            "ListTemplateType": 0,
            "Title": "About",
            "Url": "/sites/team-a/sitepages/about.aspx"
          });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Id": 2001,
          "IsDocLib": true,
          "IsExternal": false,
          "IsVisible": true,
          "ListTemplateType": 0,
          "Title": "About",
          "Url": "/sites/team-a/sitepages/about.aspx"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new navigation node to the top navigation (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1) {
        return Promise.resolve(
          {
            "Id": 2001,
            "IsDocLib": true,
            "IsExternal": false,
            "IsVisible": true,
            "ListTemplateType": 0,
            "Title": "About",
            "Url": "/sites/team-a/sitepages/about.aspx"
          });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Id": 2001,
          "IsDocLib": true,
          "IsExternal": false,
          "IsVisible": true,
          "ListTemplateType": 0,
          "Title": "About",
          "Url": "/sites/team-a/sitepages/about.aspx"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new navigation node pointing to an external URL to the quick launch', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/quicklaunch`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          Title: 'About us',
          Url: 'https://contoso.com/about-us',
          IsExternal: true
        })) {
        return Promise.resolve(
          {
            "Id": 2001,
            "IsDocLib": true,
            "IsExternal": true,
            "IsVisible": true,
            "ListTemplateType": 0,
            "Title": "About us",
            "Url": "https://contoso.com/about-us"
          });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch', title: 'About us', url: 'https://contoso.com/about-us', isExternal: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Id": 2001,
          "IsDocLib": true,
          "IsExternal": true,
          "IsVisible": true,
          "ListTemplateType": 0,
          "Title": "About us",
          "Url": "https://contoso.com/about-us"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new navigation node below an existing node', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/GetNodeById(1000)/Children`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          Title: 'About',
          Url: '/sites/team-a/sitepages/about.aspx',
          IsExternal: false
        })) {
        return Promise.resolve(
          {
            "Id": 2001,
            "IsDocLib": true,
            "IsExternal": false,
            "IsVisible": true,
            "ListTemplateType": 0,
            "Title": "About",
            "Url": "/sites/team-a/sitepages/about.aspx"
          });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', parentNodeId: 1000, title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Id": 2001,
          "IsDocLib": true,
          "IsExternal": false,
          "IsVisible": true,
          "ListTemplateType": 0,
          "Title": "About",
          "Url": "/sites/team-a/sitepages/about.aspx"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1) {
        return Promise.reject({ error: 'An error has occurred' });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } } as any, (err?: any) => {
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
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } } as any, (err?: any) => {
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
    const actual = await command.validate({ options: { webUrl: 'invalid', location: 'TopNavigationBar', title: 'About', url: '/sites/team-s/sitepages/about.aspx' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified parentNodeId is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: 'About', url: '/sites/team-s/sitepages/about.aspx', parentNodeId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified location is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'invalid', title: 'About', url: '/sites/team-s/sitepages/about.aspx' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is QuickLaunch and all required properties are present', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and the link is external', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx', isExternal: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is QuickLaunch and the link is external', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch', title: 'About', url: '/sites/team-a/sitepages/about.aspx', isExternal: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is not specified but parentNodeId is', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', parentNodeId: 2000, title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});