import assert from 'assert';
import sinon from 'sinon';
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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './tenant-site-rename.js';
import { settingsNames } from '../../../../settingsNames.js';
import { timersUtil } from '../../../../utils/timersUtil.js';

describe(commands.TENANT_SITE_RENAME, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(timersUtil, 'setTimeout').resolves();
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    const futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);
    sinon.stub(spo, 'getRequestDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: futureDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' });

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      spo.getRequestDigest,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SITE_RENAME);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a site rename job using new url parameter', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {
        return {
          "Option": 0,
          "Reserve": null,
          "OperationId": "00000000-0000-0000-0000-000000000000",
          "SkipGestures": "",
          "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1",
          "TargetSiteTitle": null,
          "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-renamed",
          "ErrorCode": 0,
          "ErrorDescription": null,
          "JobId": "76b7d932-1fb5-4fca-a336-fcceb03e157b",
          "JobState": "Success",
          "ParentId": "00000000-0000-0000-0000-000000000000",
          "SiteId": "18f8cd3b-c000-0000-0000-48bfd83e50c1",
          "TriggeredBy": "user@contoso.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/site1', newUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', verbose: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('creates a site rename job - json output', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {
        return {
          "Option": 0,
          "Reserve": null,
          "OperationId": "00000000-0000-0000-0000-000000000000",
          "SkipGestures": "",
          "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1",
          "TargetSiteTitle": null,
          "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-renamed",
          "ErrorCode": 0,
          "ErrorDescription": null,
          "JobId": "76b7d932-1fb5-4fca-a336-fcceb03e157b",
          "JobState": "Success",
          "ParentId": "00000000-0000-0000-0000-000000000000",
          "SiteId": "18f8cd3b-c000-0000-0000-48bfd83e50c1",
          "TriggeredBy": "user@contoso.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { output: 'json', url: 'https://contoso.sharepoint.com/sites/site1', newUrl: 'https://contoso.sharepoint.com/sites/site1-renamed' } });
    assert(loggerLogSpy.calledWith({
      "Option": 0,
      "Reserve": null,
      "OperationId": "00000000-0000-0000-0000-000000000000",
      "SkipGestures": "",
      "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1",
      "TargetSiteTitle": null,
      "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-renamed",
      "ErrorCode": 0,
      "ErrorDescription": null,
      "JobId": "76b7d932-1fb5-4fca-a336-fcceb03e157b",
      "JobState": "Success",
      "ParentId": "00000000-0000-0000-0000-000000000000",
      "SiteId": "18f8cd3b-c000-0000-0000-48bfd83e50c1",
      "TriggeredBy": "user@contoso.onmicrosoft.com"
    }));
  });

  it('creates a site rename job using new url parameter - suppressMarketplaceAppCheck flag', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1
        && opts.data.Option === 8) {
        return {
          "Option": 8,
          "Reserve": null,
          "OperationId": "00000000-0000-0000-0000-000000000000",
          "SkipGestures": "",
          "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1",
          "TargetSiteTitle": null,
          "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-renamed",
          "ErrorCode": 0,
          "ErrorDescription": null,
          "JobId": "76b7d932-1fb5-4fca-a336-fcceb03e157b",
          "JobState": "Success",
          "ParentId": "00000000-0000-0000-0000-000000000000",
          "SiteId": "18f8cd3b-c000-0000-0000-48bfd83e50c1",
          "TriggeredBy": "user@contoso.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/site1', newUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', suppressMarketplaceAppCheck: true, verbose: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('creates a site rename job using new url parameter - suppressWorkflow2013Check flag', async () => {

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1
        && opts.data.Option === 16) {
        return {
          "Option": 16,
          "Reserve": null,
          "OperationId": "00000000-0000-0000-0000-000000000000",
          "SkipGestures": "",
          "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1",
          "TargetSiteTitle": null,
          "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-renamed",
          "ErrorCode": 0,
          "ErrorDescription": null,
          "JobId": "76b7d932-1fb5-4fca-a336-fcceb03e157b",
          "JobState": "Success",
          "ParentId": "00000000-0000-0000-0000-000000000000",
          "SiteId": "18f8cd3b-c000-0000-0000-48bfd83e50c1",
          "TriggeredBy": "user@contoso.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/site1', newUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', suppressWorkflow2013Check: true, verbose: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('creates a site rename job using new url parameter - both supress flags', async () => {

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1
        && opts.data.Option === 24) {
        return {
          "Option": 24,
          "Reserve": null,
          "OperationId": "00000000-0000-0000-0000-000000000000",
          "SkipGestures": "",
          "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1",
          "TargetSiteTitle": "RenamedSite",
          "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-renamed",
          "ErrorCode": 0,
          "ErrorDescription": null,
          "JobId": "76b7d932-1fb5-4fca-a336-fcceb03e157b",
          "JobState": "Success",
          "ParentId": "00000000-0000-0000-0000-000000000000",
          "SiteId": "18f8cd3b-c000-0000-0000-48bfd83e50c1",
          "TriggeredBy": "user@contoso.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/site1', newUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', newTitle: "RenamedSite", suppressWorkflow2013Check: true, suppressMarketplaceAppCheck: true, verbose: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('creates a site rename job - wait for completion', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {
        return {
          "Option": 0,
          "Reserve": null,
          "OperationId": "00000000-0000-0000-0000-000000000000",
          "SkipGestures": "",
          "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1",
          "TargetSiteTitle": null,
          "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-renamed",
          "ErrorCode": 0,
          "ErrorDescription": null,
          "JobId": "76b7d932-1fb5-4fca-a336-fcceb03e157b",
          "JobState": "NotStarted",
          "ParentId": "00000000-0000-0000-0000-000000000000",
          "SiteId": "18f8cd3b-c000-0000-0000-48bfd83e50c1",
          "TriggeredBy": "user@contoso.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/SiteRenameJobs/GetJobsBySiteUrl') > -1 &&
        opts.headers &&
        opts.headers['X-AttemptNumber'] &&
        parseInt(opts.headers['X-AttemptNumber'] as string) <= 1) {
        return {
          "odata.metadata": "https://contoso-admin.sharepoint.com/_api/$metadata#SP.ApiData.SiteRenameJobEntityDatas",
          "value":
            [{
              "odata.type": "Microsoft.Online.SharePoint.Onboarding.RestService.Service.SiteRenameJob",
              "odata.id": "https://contoso-admin.sharepoint.com/_api/Microsoft.Online.SharePoint.Onboarding.RestService.Service.SiteRenameJobc416c883-a2b5-465b-b595-683500e83c72",
              "odata.editLink": "Microsoft.Online.SharePoint.Onboarding.RestService.Service.SiteRenameJobc416c883-a2b5-465b-b595-683500e83c72",
              "Option": 0,
              "Reserve": null,
              "OperationId": "00000000-0000-0000-0000-000000000000",
              "SkipGestures": null,
              "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1",
              "TargetSiteTitle": null,
              "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site2",
              "ErrorCode": 0,
              "ErrorDescription": null,
              "JobId": "3080d202-27a5-4392-8139-e94d2379c109",
              "JobState": "NotStarted",
              "ParentId": "00000000-0000-0000-0000-000000000000",
              "SiteId": "63f68a25-460d-4626-bf79-aca4bb158ca8",
              "TriggeredBy": "user@contoso.onmicrosoft.com"
            }]
        };
      }
      else if ((opts.url as string).indexOf('/_api/SiteRenameJobs/GetJobsBySiteUrl') > -1 &&
        opts.headers &&
        opts.headers['X-AttemptNumber'] &&
        parseInt(opts.headers['X-AttemptNumber'] as string) > 1) {
        return {
          "odata.metadata": "https://contoso-admin.sharepoint.com/_api/$metadata#SP.ApiData.SiteRenameJobEntityDatas",
          "value":
            [{
              "odata.type": "Microsoft.Online.SharePoint.Onboarding.RestService.Service.SiteRenameJob",
              "odata.id": "https://contoso-admin.sharepoint.com/_api/Microsoft.Online.SharePoint.Onboarding.RestService.Service.SiteRenameJobc416c883-a2b5-465b-b595-683500e83c72",
              "odata.editLink": "Microsoft.Online.SharePoint.Onboarding.RestService.Service.SiteRenameJobc416c883-a2b5-465b-b595-683500e83c72",
              "Option": 0, "Reserve": null,
              "OperationId": "00000000-0000-0000-0000-000000000000",
              "SkipGestures": null,
              "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1",
              "TargetSiteTitle": null,
              "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site2",
              "ErrorCode": 0,
              "ErrorDescription": null,
              "JobId": "3080d202-27a5-4392-8139-e94d2379c109",
              "JobState": "Success", "ParentId": "00000000-0000-0000-0000-000000000000",
              "SiteId": "63f68a25-460d-4626-bf79-aca4bb158ca8",
              "TriggeredBy": "user@contoso.onmicrosoft.com"
            }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/site1', newUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', wait: true, debug: true, verbose: true } } as any);
    assert(loggerLogToStderrSpy.called);
  });

  it('handles API error - delayed failure - valid response', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {
        return {
          "Option": 0,
          "Reserve": null,
          "OperationId": "00000000-0000-0000-0000-000000000000",
          "SkipGestures": "",
          "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1-reject",
          "TargetSiteTitle": null,
          "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-reject-renamed",
          "ErrorCode": 0,
          "ErrorDescription": null,
          "JobId": "76b7d932-1fb5-4fca-a336-fcceb03e157b",
          "JobState": "NotStarted",
          "ParentId": "00000000-0000-0000-0000-000000000000",
          "SiteId": "18f8cd3b-c000-0000-0000-48bfd83e50c1",
          "TriggeredBy": "user@contoso.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/SiteRenameJobs/GetJobsBySiteUrl') > -1) {
        return {
          "odata.metadata": "https://contoso-admin.sharepoint.com/_api/$metadata#SP.ApiData.SiteRenameJobEntityDatas",
          "value":
            [{
              "odata.type": "Microsoft.Online.SharePoint.Onboarding.RestService.Service.SiteRenameJob",
              "odata.id": "https://contoso-admin.sharepoint.com/_api/Microsoft.Online.SharePoint.Onboarding.RestService.Service.SiteRenameJobc416c883-a2b5-465b-b595-683500e83c72",
              "odata.editLink": "Microsoft.Online.SharePoint.Onboarding.RestService.Service.SiteRenameJobc416c883-a2b5-465b-b595-683500e83c72",
              "Option": 0, "Reserve": null,
              "OperationId": "00000000-0000-0000-0000-000000000000",
              "SkipGestures": null,
              "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1-reject",
              "TargetSiteTitle": null,
              "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-reject-renamed",
              "ErrorCode": 123,
              "ErrorDescription": "An error has occurred",
              "JobId": "3080d202-27a5-4392-8139-e94d2379c109",
              "JobState": "Failed", "ParentId": "00000000-0000-0000-0000-000000000000",
              "SiteId": "63f68a25-460d-4626-bf79-aca4bb158ca8",
              "TriggeredBy": "user@contoso.onmicrosoft.com"
            }]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        url: "https://contoso.sharepoint.com/sites/site1-reject",
        newUrl: "https://contoso.sharepoint.com/sites/site1-reject-renamed",
        wait: true,
        verbose: true
      }
    } as any), new CommandError("An error has occurred"));
  });

  it('handles API error - delayed failure - service error', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {
        return {
          "Option": 0,
          "Reserve": null,
          "OperationId": "00000000-0000-0000-0000-000000000000",
          "SkipGestures": "",
          "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1-reject",
          "TargetSiteTitle": null,
          "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-reject-renamed",
          "ErrorCode": 0,
          "ErrorDescription": null,
          "JobId": "76b7d932-1fb5-4fca-a336-fcceb03e157b",
          "JobState": "NotStarted",
          "ParentId": "00000000-0000-0000-0000-000000000000",
          "SiteId": "18f8cd3b-c000-0000-0000-48bfd83e50c1",
          "TriggeredBy": "user@contoso.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(() => {
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        url: "https://contoso.sharepoint.com/sites/site1-reject",
        newUrl: "https://contoso.sharepoint.com/sites/site1-reject-renamed",
        wait: true,
        verbose: true
      }
    } as any), new CommandError("Invalid request"));
  });

  it('handles API error - immediate failure on creation', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {

        return {
          "Option": 0,
          "Reserve": null,
          "OperationId": "00000000-0000-0000-0000-000000000000",
          "SkipGestures": "",
          "SourceSiteUrl": "https://contoso.sharepoint.com/sites/site1",
          "TargetSiteTitle": null,
          "TargetSiteUrl": "https://contoso.sharepoint.com/sites/site1-renamed",
          "ErrorCode": 0,
          "ErrorDescription": "An error has occurred",
          "JobId": "76b7d932-1fb5-4fca-a336-fcceb03e157b",
          "JobState": "Error",
          "ParentId": "00000000-0000-0000-0000-000000000000",
          "SiteId": "18f8cd3b-c000-0000-0000-48bfd83e50c1",
          "TriggeredBy": "user@contoso.onmicrosoft.com"
        };

      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        url: "https://contoso.sharepoint.com/sites/old",
        newUrl: "https://contoso.sharepoint.com/sites/new",
        wait: true
      }
    } as any), new CommandError("An error has occurred"));
  });

  it('accepts newUrl parameter', async () => {
    const actual = await command.validate({ options: { url: "https://contoso.sharepoint.com/", newUrl: "https://contoso.sharepoint.com/sites/new" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts both newUrl and newTitle', async () => {
    const actual = await command.validate({ options: { url: "https://contoso.sharepoint.com/", newUrl: "https://contoso.sharepoint.com/sites/new", newTitle: "New Site" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts suppressMarketplaceAppCheck flag', async () => {
    const actual = await command.validate({ options: { url: "https://contoso.sharepoint.com/", newUrl: "https://contoso.sharepoint.com/sites/new", newTitle: "New Site", suppressMarketplaceAppCheck: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts suppressWorkflow2013Check flag', async () => {
    const actual = await command.validate({ options: { url: "https://contoso.sharepoint.com/", newUrl: "https://contoso.sharepoint.com/sites/new", newTitle: "New Site", suppressWorkflow2013Check: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts wait flag', async () => {
    const actual = await command.validate({ options: { url: "https://contoso.sharepoint.com/", newUrl: "https://contoso.sharepoint.com/sites/new", newTitle: "New Site", wait: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects missing newUrl', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { url: "https://contoso.sharepoint.com", newTitle: "New Site" } }, commandInfo);
    assert.strictEqual(actual, `Required option newUrl not specified`);
  });

  it('rejects when newUrl is the same as url', async () => {
    const actual = await command.validate({ options: { url: "https://contoso.sharepoint.com/sites/target", newUrl: "https://contoso.sharepoint.com/sites/target" } }, commandInfo);
    assert.strictEqual(actual, `The new URL cannot be the same as the target URL.`);
  });
});
