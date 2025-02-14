import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './site-get.js';

describe(commands.SITE_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
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
    sinon.restore();
    auth.connection.active = false;
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
        "DecodedUrl": "https://contoso.sharepoint.com/sites/sales"
      },
      "PrimaryUri": "https://contoso.sharepoint.com/sites/sales",
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
      "Url": "https://contoso.sharepoint.com/sites/project-x"
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/site`) > -1) {
        return siteProperties;
      }

      throw 'Invalid request';
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
        "DecodedUrl": "https://contoso.sharepoint.com/sites/sales"
      },
      "PrimaryUri": "https://contoso.sharepoint.com/sites/sales",
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
      "Url": "https://contoso.sharepoint.com/sites/project-x"
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/site`) > -1) {
        return siteProperties;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert(loggerLogSpy.calledWith(siteProperties));
  });

  it('correctly handles error when getting information for a site that doesn\'t exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/site') > -1) {
        throw new Error("404 - \"404 FILE NOT FOUND\"");
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/project-x' } } as any), new CommandError('404 - "404 FILE NOT FOUND"'));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'foo' });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, true);
  });
});
