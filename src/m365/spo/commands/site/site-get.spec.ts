import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./site-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SITE_GET, () => {
  let log: any[];
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
    assert.strictEqual(command.name.startsWith(commands.SITE_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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
      if ((opts.url as string).indexOf(`/_api/site`) > -1) {
        return Promise.resolve(siteProperties);
      }

      return Promise.reject('Invalid request');
    });

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
      if ((opts.url as string).indexOf(`/_api/site`) > -1) {
        return Promise.resolve(siteProperties);
      }

      return Promise.reject('Invalid request');
    });

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
      if ((opts.url as string).indexOf('/_api/site') > -1) {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/project-x' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('404 - "404 FILE NOT FOUND"')));
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

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });
});