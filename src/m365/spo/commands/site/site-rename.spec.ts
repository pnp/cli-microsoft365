import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./site-rename');

describe(commands.SITE_RENAME, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    const futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);
    sinon.stub(spo, 'getRequestDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: futureDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      spo.getRequestDigest
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      global.setTimeout
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_RENAME), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a site rename job using new url parameter', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', verbose: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a site rename job - json output', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { output: 'json', siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed' } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a site rename job using new url parameter - suppressMarketplaceAppCheck flag', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1
        && opts.data.Option === 8) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', suppressMarketplaceAppCheck: true, verbose: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a site rename job using new url parameter - suppressWorkflow2013Check flag', (done) => {

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1
        && opts.data.Option === 16) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', suppressWorkflow2013Check: true, verbose: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a site rename job using new url parameter - both supress flags', (done) => {

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1
        && opts.data.Option === 24) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', newSiteTitle: "RenamedSite", suppressWorkflow2013Check: true, suppressMarketplaceAppCheck: true, verbose: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a site rename job - wait for completion', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/SiteRenameJobs/GetJobsBySiteUrl') > -1 &&
        opts.headers &&
        opts.headers['X-AttemptNumber'] &&
        parseInt(opts.headers['X-AttemptNumber'] as string) <= 1) {
        return Promise.resolve(
          {
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
          }
        );
      }
      else if ((opts.url as string).indexOf('/_api/SiteRenameJobs/GetJobsBySiteUrl') > -1 &&
        opts.headers &&
        opts.headers['X-AttemptNumber'] &&
        parseInt(opts.headers['X-AttemptNumber'] as string) > 1) {
        return Promise.resolve(
          {
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
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', wait: true, debug: true, verbose: true } } as any, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles API error - delayed failure - valid response', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/SiteRenameJobs/GetJobsBySiteUrl') > -1) {
        return Promise.resolve(
          {
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
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/site1-reject",
        newSiteUrl: "https://contoso.sharepoint.com/sites/site1-reject-renamed",
        wait: true,
        verbose: true
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles API error - delayed failure - service error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/site1-reject",
        newSiteUrl: "https://contoso.sharepoint.com/sites/site1-reject-renamed",
        wait: true,
        verbose: true
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Invalid request")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles API error - immediate failure on creation', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SiteRenameJobs?api-version=1.4.7`) > -1) {

        return Promise.resolve({
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
        });

      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/old",
        newSiteUrl: "https://contoso.sharepoint.com/sites/new",
        wait: true
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('accepts newSiteUrl parameter', async () => {
    const actual = await command.validate({ options: { siteUrl: "https://contoso.sharepoint.com/", newSiteUrl: "https://contoso.sharepoint.com/sites/new" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts both newSiteUrl and newSiteTitle', async () => {
    const actual = await command.validate({ options: { siteUrl: "https://contoso.sharepoint.com/", newSiteUrl: "https://contoso.sharepoint.com/sites/new", newSiteTitle: "New Site" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts suppressMarketplaceAppCheck flag', async () => {
    const actual = await command.validate({ options: { siteUrl: "https://contoso.sharepoint.com/", newSiteUrl: "https://contoso.sharepoint.com/sites/new", newSiteTitle: "New Site", suppressMarketplaceAppCheck: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts suppressWorkflow2013Check flag', async () => {
    const actual = await command.validate({ options: { siteUrl: "https://contoso.sharepoint.com/", newSiteUrl: "https://contoso.sharepoint.com/sites/new", newSiteTitle: "New Site", suppressWorkflow2013Check: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts wait flag', async () => {
    const actual = await command.validate({ options: { siteUrl: "https://contoso.sharepoint.com/", newSiteUrl: "https://contoso.sharepoint.com/sites/new", newSiteTitle: "New Site", wait: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects missing newSiteUrl', async () => {
    const actual = await command.validate({ options: { siteUrl: "https://contoso.sharepoint.com", newSiteTitle: "New Site" } }, commandInfo);
    assert.strictEqual(actual, `Required option newSiteUrl not specified`);
  });

  it('rejects when newSiteUrl is the same as siteUrl', async () => {
    const actual = await command.validate({ options: { siteUrl: "https://contoso.sharepoint.com/sites/target", newSiteUrl: "https://contoso.sharepoint.com/sites/target" } }, commandInfo);
    assert.strictEqual(actual, `The new URL cannot be the same as the target URL.`);
  });
});