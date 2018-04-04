import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./hubsite-data-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.HUBSITE_DATA_GET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.HUBSITE_DATA_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.HUBSITE_DATA_GET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/HubSiteData(false)`) > -1) {
        return Promise.resolve({
          value: JSON.stringify({
            "themeKey": null,
            "name": "CommunicationSite",
            "url": "https://contoso.sharepoint.com/sites/Sales",
            "logoUrl": "http://contoso.com/__siteIcon__.jpg",
            "usesMetadataNavigation": false,
            "navigation": []
          })
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "themeKey": null,
          "name": "CommunicationSite",
          "url": "https://contoso.sharepoint.com/sites/Sales",
          "logoUrl": "http://contoso.com/__siteIcon__.jpg",
          "usesMetadataNavigation": false,
          "navigation": []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site with forced refresh', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/HubSiteData(true)`) > -1) {
        return Promise.resolve({
          value: JSON.stringify({
            "themeKey": null,
            "name": "CommunicationSite",
            "url": "https://contoso.sharepoint.com/sites/Sales",
            "logoUrl": "http://contoso.com/__siteIcon__.jpg",
            "usesMetadataNavigation": false,
            "navigation": []
          })
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X', forceRefresh: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "themeKey": null,
          "name": "CommunicationSite",
          "url": "https://contoso.sharepoint.com/sites/Sales",
          "logoUrl": "http://contoso.com/__siteIcon__.jpg",
          "usesMetadataNavigation": false,
          "navigation": []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/HubSiteData`) > -1) {
        return Promise.resolve({
          value: JSON.stringify({
            "themeKey": null,
            "name": "CommunicationSite",
            "url": "https://contoso.sharepoint.com/sites/Sales",
            "logoUrl": "http://contoso.com/__siteIcon__.jpg",
            "usesMetadataNavigation": false,
            "navigation": []
          })
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "themeKey": null,
          "name": "CommunicationSite",
          "url": "https://contoso.sharepoint.com/sites/Sales",
          "logoUrl": "http://contoso.com/__siteIcon__.jpg",
          "usesMetadataNavigation": false,
          "navigation": []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/HubSiteData`) > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/HubSiteData`) > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(`https://contoso.sharepoint.com/sites/Project-X is not connected to a hub site and is not a hub site itself`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when hub site not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-1, Microsoft.SharePoint.Client.ResourceNotFoundException",
            "message": {
              "lang": "en-US",
              "value": "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
            }
          }
        }
      });
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying forceRefresh', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--forceRefresh') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webUrl: 'Invalid' } });
    assert.notEqual(actual, true);
  });

  it('passed validation if webUrl is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webUrl: 'https://contoso.sharepoint.com' } });
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
    assert(find.calledWith(commands.HUBSITE_DATA_GET));
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

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/Sales', confirm: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});