import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./site-rename');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.SITE_RENAME, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return { FormDigestValue: 'abc' }; });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    let futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: futureDate.toISOString() }); });

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
  })

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post,
      (command as any).ensureFormDigest
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
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

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
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

    cmdInstance.action({ options: { output: 'json', siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
        && opts.body.Option === 8) {
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

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', suppressMarketplaceAppCheck: true, verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
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
        && opts.body.Option === 16) {
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

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', suppressWorkflow2013Check: true, verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
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
        && opts.body.Option === 24) {
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

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', newSiteTitle: "RenamedSite", suppressWorkflow2013Check: true, suppressMarketplaceAppCheck: true, verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
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
        parseInt(opts.headers['X-AttemptNumber']) <= 1) {
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
        parseInt(opts.headers['X-AttemptNumber']) > 1) {
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

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/site1', newSiteUrl: 'https://contoso.sharepoint.com/sites/site1-renamed', wait: true, debug: true, verbose: true } }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
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

    cmdInstance.action({
      options: {
        siteUrl: "http://contoso.sharepoint.com/sites/site1-reject",
        newSiteUrl: "http://contoso.sharepoint.com/sites/site1-reject-renamed",
        wait: true,
        verbose: true
      }
    }, (err?: any) => {
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

    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        siteUrl: "http://contoso.sharepoint.com/sites/site1-reject",
        newSiteUrl: "http://contoso.sharepoint.com/sites/site1-reject-renamed",
        wait: true,
        verbose: true
      }
    }, (err?: any) => {
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

    cmdInstance.action({
      options: {
        siteUrl: "http://contoso.sharepoint.com/sites/old",
        newSiteUrl: "http://contoso.sharepoint.com/sites/new",
        wait: true
      }
    }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('accepts newSiteUrl parameter', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: "http://contoso.sharepoint.com/", newSiteUrl: "http://contoso.sharepoint.com/sites/new" } });
    assert.strictEqual(actual, true);
  });

  it('accepts both newSiteUrl and newSiteTitle', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: "http://contoso.sharepoint.com/", newSiteUrl: "http://contoso.sharepoint.com/sites/new", newSiteTitle: "New Site" } });
    assert.strictEqual(actual, true);
  });

  it('accepts suppressMarketplaceAppCheck flag', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: "http://contoso.sharepoint.com/", newSiteUrl: "http://contoso.sharepoint.com/sites/new", newSiteTitle: "New Site", suppressMarketplaceAppCheck: true } });
    assert.strictEqual(actual, true);
  });

  it('accepts suppressWorkflow2013Check flag', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: "http://contoso.sharepoint.com/", newSiteUrl: "http://contoso.sharepoint.com/sites/new", newSiteTitle: "New Site", suppressWorkflow2013Check: true } });
    assert.strictEqual(actual, true);
  });

  it('accepts wait flag', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: "http://contoso.sharepoint.com/", newSiteUrl: "http://contoso.sharepoint.com/sites/new", newSiteTitle: "New Site", wait: true } });
    assert.strictEqual(actual, true);
  });

  it('rejects missing newSiteUrl', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: "http://contoso.sharepoint.com", newSiteTitle: "New Site" } });
    assert.strictEqual(actual, `A new url must be provided.`);
  });

  it('rejects when newSiteUrl is the same as siteUrl', () => {
    const actual = (command.validate() as CommandValidate)({ options: { siteUrl: "http://contoso.sharepoint.com/sites/target", newSiteUrl: "http://contoso.sharepoint.com/sites/target" } });
    assert.strictEqual(actual, `The new URL cannot be the same as the target URL.`);
  });
});