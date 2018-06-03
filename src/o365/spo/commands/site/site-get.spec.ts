import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./site-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.SITE_GET, () => {
  let vorpal: Vorpal;
  let log: any[];
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
    assert.equal(command.name.startsWith(commands.SITE_GET), true);
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
        assert.equal(telemetry.name, commands.SITE_GET);
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
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves the information for the specified site', (done) => {
    const siteProperties = {
      "AllowCreateDeclarativeWorkflow": true,
      "AllowDesigner": true,
      "AllowMasterPageEditing": false,
      "AllowRevertFromTemplate": false,
      "AllowSaveDeclarativeWorkflowAsTemplate": true,
      "AllowSavePublishDeclarativeWorkflow": true,
      "AllowSelfServiceUpgrade": true,
      "AllowSelfServiceUpgradeEvaluation": true,
      "AuditLogTrimmingRetention": 90,
      "Classification": "",
      "CompatibilityLevel": 15,
      "CurrentChangeToken": {
        "StringValue": "1;1;1a70e568-d286-4ad1-b036-734ff8667915;636527399616270000;66855110"
      },
      "DisableAppViews": false,
      "DisableCompanyWideSharingLinks": false,
      "DisableFlows": false,
      "ExternalSharingTipsEnabled": false,
      "GeoLocation": "EUR",
      "GroupId": "7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091",
      "HubSiteId": "00000000-0000-0000-0000-000000000000",
      "Id": "1a70e568-d286-4ad1-b036-734ff8667915",
      "IsHubSite": false,
      "LockIssue": null,
      "MaxItemsPerThrottledOperation": 5000,
      "NeedsB2BUpgrade": false,
      "ResourcePath": {
        "DecodedUrl": "https://m365x324230.sharepoint.com/sites/sales"
      },
      "PrimaryUri": "https://m365x324230.sharepoint.com/sites/sales",
      "ReadOnly": false,
      "RequiredDesignerVersion": "15.0.0.0",
      "SandboxedCodeActivationCapability": 2,
      "ServerRelativeUrl": "/sites/sales",
      "ShareByEmailEnabled": true,
      "ShareByLinkEnabled": false,
      "ShowUrlStructure": false,
      "TrimAuditLog": true,
      "UIVersionConfigurationEnabled": false,
      "UpgradeReminderDate": "1899-12-30T00:00:00",
      "UpgradeScheduled": false,
      "UpgradeScheduledDate": "1753-01-01T00:00:00",
      "Upgrading": false,
      "Url": "https://m365x324230.sharepoint.com/sites/project-x"
    };
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/site`) > -1 &&
        opts.headers.authorization.indexOf('Bearer ') === 0) {
        return Promise.resolve(siteProperties);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(siteProperties));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves the information for the specified site (debug)', (done) => {
    const siteProperties = {
      "AllowCreateDeclarativeWorkflow": true,
      "AllowDesigner": true,
      "AllowMasterPageEditing": false,
      "AllowRevertFromTemplate": false,
      "AllowSaveDeclarativeWorkflowAsTemplate": true,
      "AllowSavePublishDeclarativeWorkflow": true,
      "AllowSelfServiceUpgrade": true,
      "AllowSelfServiceUpgradeEvaluation": true,
      "AuditLogTrimmingRetention": 90,
      "Classification": "",
      "CompatibilityLevel": 15,
      "CurrentChangeToken": {
        "StringValue": "1;1;1a70e568-d286-4ad1-b036-734ff8667915;636527399616270000;66855110"
      },
      "DisableAppViews": false,
      "DisableCompanyWideSharingLinks": false,
      "DisableFlows": false,
      "ExternalSharingTipsEnabled": false,
      "GeoLocation": "EUR",
      "GroupId": "7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091",
      "HubSiteId": "00000000-0000-0000-0000-000000000000",
      "Id": "1a70e568-d286-4ad1-b036-734ff8667915",
      "IsHubSite": false,
      "LockIssue": null,
      "MaxItemsPerThrottledOperation": 5000,
      "NeedsB2BUpgrade": false,
      "ResourcePath": {
        "DecodedUrl": "https://m365x324230.sharepoint.com/sites/sales"
      },
      "PrimaryUri": "https://m365x324230.sharepoint.com/sites/sales",
      "ReadOnly": false,
      "RequiredDesignerVersion": "15.0.0.0",
      "SandboxedCodeActivationCapability": 2,
      "ServerRelativeUrl": "/sites/sales",
      "ShareByEmailEnabled": true,
      "ShareByLinkEnabled": false,
      "ShowUrlStructure": false,
      "TrimAuditLog": true,
      "UIVersionConfigurationEnabled": false,
      "UpgradeReminderDate": "1899-12-30T00:00:00",
      "UpgradeScheduled": false,
      "UpgradeScheduledDate": "1753-01-01T00:00:00",
      "Upgrading": false,
      "Url": "https://m365x324230.sharepoint.com/sites/project-x"
    };
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/site`) > -1 &&
        opts.headers.authorization.indexOf('Bearer ') === 0) {
        return Promise.resolve(siteProperties);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(siteProperties));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when getting information for a site that doesn\'t exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/site') > -1) {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/project-x' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('404 - "404 FILE NOT FOUND"')));
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

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<url>') > -1) {
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
    const actual = (command.validate() as CommandValidate)({ options: { url: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com' } });
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
    assert(find.calledWith(commands.SITE_GET));
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
    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/project-x' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});