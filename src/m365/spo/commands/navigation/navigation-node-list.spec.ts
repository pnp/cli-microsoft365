import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./navigation-node-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.NAVIGATION_NODE_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{ "Id": 2003, "Title": "Node 1", "Url": "/sites/team-a/SitePages/page1.aspx" }, { "Id": 2004, "Title": "Node 2", "Url": "/sites/team-a/SitePages/page2.aspx" }]));
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

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{ "Id": 2003, "Title": "Node 1", "Url": "/sites/team-a/SitePages/page1.aspx" }, { "Id": 2004, "Title": "Node 2", "Url": "/sites/team-a/SitePages/page2.aspx" }]));
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } }, (err?: any) => {
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid', location: 'TopNavigationBar', title: 'About', url: '/sites/team-s/sitepages/about.aspx' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified location is not valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'invalid', title: 'About', url: '/sites/team-s/sitepages/about.aspx' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is QuickLaunch and all required properties are present', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' } });
    assert.strictEqual(actual, true);
  });
});