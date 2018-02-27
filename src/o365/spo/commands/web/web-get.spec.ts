import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./web-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.WEB_GET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
  let stubAuth: any = () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return { FormDigestValue: 'abc' }; });
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
      auth.restoreAuth,
      request.get
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.WEB_GET), true);
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
        assert.equal(telemetry.name, commands.WEB_GET);
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
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves site information', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "AllowRssFeeds": false,
            "AlternateCssUrl": null,
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "Configuration": 0,
            "Created": null,
            "CurrentChangeToken": null,
            "CustomMasterUrl": null,
            "Description": null,
            "DesignPackageId": null,
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": false,
            "HorizontalQuickLaunch": false,
            "Id": "d8d179c7-f459-4f90-b592-14b08e84accb",
            "IsMultilingual": false,
            "Language": 1033,
            "LastItemModifiedDate": null,
            "LastItemUserModifiedDate": null,
            "MasterUrl": null,
            "NoCrawl": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": null,
            "QuickLaunchEnabled": false,
            "RecycleBinEnabled": false,
            "ServerRelativeUrl": null,
            "SiteLogoUrl": null,
            "SyndicationEnabled": false,
            "Title": "Subsite",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/subsite",
            "WebTemplate": "STS",
             }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          value: [{
            AllowRssFeeds: false,
            AlternateCssUrl: null,
            AppInstanceId: "00000000-0000-0000-0000-000000000000",
            Configuration: 0,
            Created: null,
            CurrentChangeToken: null,
            CustomMasterUrl: null,
            Description: null,
            DesignPackageId: null,
            DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false,
            EnableMinimalDownload: false,
            HorizontalQuickLaunch: false,
            Id: "d8d179c7-f459-4f90-b592-14b08e84accb",
            IsMultilingual: false,
            Language: 1033,
            LastItemModifiedDate: null,
            LastItemUserModifiedDate: null,
            MasterUrl: null,
            NoCrawl: false,
            OverwriteTranslationsOnChange: false,
            ResourcePath: null,
            QuickLaunchEnabled: false,
            RecycleBinEnabled: false,
            ServerRelativeUrl: null,
            SiteLogoUrl: null,
            SyndicationEnabled: false,
            Title: "Subsite",
            TreeViewEnabled: false,
            UIVersion: 15,
            UIVersionConfigurationEnabled: false,
            Url: "https://contoso.sharepoint.com/subsite",
            WebTemplate: "STS",
            }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('retrieves all site information with output option text', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "AllowRssFeeds": false,
            "AlternateCssUrl": null,
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "Configuration": 0,
            "Created": null,
            "CurrentChangeToken": null,
            "CustomMasterUrl": null,
            "Description": null,
            "DesignPackageId": null,
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": false,
            "HorizontalQuickLaunch": false,
            "Id": "d8d179c7-f459-4f90-b592-14b08e84accb",
            "IsMultilingual": false,
            "Language": 1033,
            "LastItemModifiedDate": null,
            "LastItemUserModifiedDate": null,
            "MasterUrl": null,
            "NoCrawl": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": null,
            "QuickLaunchEnabled": false,
            "RecycleBinEnabled": false,
            "ServerRelativeUrl": null,
            "SiteLogoUrl": null,
            "SyndicationEnabled": false,
            "Title": "Subsite",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/subsite",
            "WebTemplate": "STS",
             }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          value: [{
            AllowRssFeeds: false,
            AlternateCssUrl: null,
            AppInstanceId: "00000000-0000-0000-0000-000000000000",
            Configuration: 0,
            Created: null,
            CurrentChangeToken: null,
            CustomMasterUrl: null,
            Description: null,
            DesignPackageId: null,
            DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false,
            EnableMinimalDownload: false,
            HorizontalQuickLaunch: false,
            Id: "d8d179c7-f459-4f90-b592-14b08e84accb",
            IsMultilingual: false,
            Language: 1033,
            LastItemModifiedDate: null,
            LastItemUserModifiedDate: null,
            MasterUrl: null,
            NoCrawl: false,
            OverwriteTranslationsOnChange: false,
            ResourcePath: null,
            QuickLaunchEnabled: false,
            RecycleBinEnabled: false,
            ServerRelativeUrl: null,
            SiteLogoUrl: null,
            SyndicationEnabled: false,
            Title: "Subsite",
            TreeViewEnabled: false,
            UIVersion: 15,
            UIVersionConfigurationEnabled: false,
            Url: "https://contoso.sharepoint.com/subsite",
            WebTemplate: "STS",
            }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('command correctly handles web get reject request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });

    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/webs') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
      }
    }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post,
          request.get
        ]);
      }
    });
  });

  it('uses correct API url when output json option is passed', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      cmdInstance.log('Test Url:');
      cmdInstance.log(opts.url);
      if (opts.url.indexOf('select123=') > -1) {
        return Promise.resolve('Correct Url1')
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        output: 'json',
        debug: false,
        webUrl: 'https://contoso.sharepoint.com',
      }
    }, () => {

      try {
        assert('Correct Url');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post,
          request.get
        ]);
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

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert(actual);
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
    assert(find.calledWith(commands.WEB_GET));
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
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        debug: false
      }
    }, () => {
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