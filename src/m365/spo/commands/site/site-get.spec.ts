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
const command: Command = require('./site-get');

describe(commands.SITE_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves the information for the specified site', async () => {
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

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert(loggerLogSpy.calledWith(siteProperties));
  });

  it('retrieves the information for the specified site (debug)', async () => {
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

    await command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert(loggerLogSpy.calledWith(siteProperties));
  });

  it('correctly handles error when getting information for a site that doesn\'t exist', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/site') > -1) {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/project-x' } } as any), new CommandError('404 - "404 FILE NOT FOUND"'));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<url>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
