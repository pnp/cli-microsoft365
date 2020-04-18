import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./navigation-node-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.NAVIGATION_NODE_ADD, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
      request.post
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
    assert.equal(command.name.startsWith(commands.NAVIGATION_NODE_ADD), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('adds new navigation node to the top navigation', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
        JSON.stringify(opts.body) === JSON.stringify({
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch', title: 'About us', url: 'https://contoso.com/about-us', isExternal: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
        JSON.stringify(opts.body) === JSON.stringify({
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', parentNodeId: 1000, title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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

  it('fails validation if webUrl not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { location: 'TopNavigationBar', title: 'About', url: '/sites/team-s/sitepages/about.aspx' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid', location: 'TopNavigationBar', title: 'About', url: '/sites/team-s/sitepages/about.aspx' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the specified parentNodeId is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: 'About', url: '/sites/team-s/sitepages/about.aspx', parentNodeId: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if location is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: 'About', url: '/sites/team-s/sitepages/about.aspx' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if specified location is not valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'invalid', title: 'About', url: '/sites/team-s/sitepages/about.aspx' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if title is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', url: '/sites/team-s/sitepages/about.aspx' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if url is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } });
    assert.equal(actual, true);
  });

  it('passes validation when location is QuickLaunch and all required properties are present', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch', title: 'About', url: '/sites/team-a/sitepages/about.aspx' } });
    assert.equal(actual, true);
  });

  it('passes validation when location is TopNavigationBar and the link is external', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', title: 'About', url: '/sites/team-a/sitepages/about.aspx', isExternal: true } });
    assert.equal(actual, true);
  });

  it('passes validation when location is QuickLaunch and the link is external', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch', title: 'About', url: '/sites/team-a/sitepages/about.aspx', isExternal: true } });
    assert.equal(actual, true);
  });

  it('passes validation when location is not specified but parentNodeId is', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', parentNodeId: 2000, title: 'About', url: '/sites/team-a/sitepages/about.aspx' } });
    assert.equal(actual, true);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.NAVIGATION_NODE_ADD));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});