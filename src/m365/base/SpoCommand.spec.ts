import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
import { Logger } from '../../cli';
import { CommandError, CommandOption } from '../../Command';
import request from '../../request';
import Utils from '../../Utils';
import { FormDigestInfo } from '../spo/spo';
import SpoCommand from './SpoCommand';

class MockCommand extends SpoCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(): void {
  }

  public options(): CommandOption[] {
    return [
      {
        option: '--url [url]'
      },
      {
        option: '--nonProcessedUrl [nonProcessedUrl]'
      }
    ];
  }

  public validateUnknownOptionsPublic(options: any, csomObject: string, csomPropertyType: 'get' | 'set'): string | boolean {
    return this.validateUnknownOptions(options, csomObject, csomPropertyType);
  }

  public getNamesOfOptionsWithUrlsPublic(): string[] {
    return this.getNamesOfOptionsWithUrls();
  }
}

describe('SpoCommand', () => {
  let loggerLogSpy: sinon.SinonSpy;
  let logger: Logger;
  let log: string[];

  before(() => {
    auth.service.connected = true;
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    Utils.restore([
      request.get,
      request.post,
      auth.storeConnectionInfo
    ]);
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('correctly reports an error while restoring auth info', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();

    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    command.action(logger, { options: {} } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(auth.restoreAuth);
      }
    });
  });

  it('doesn\'t execute command when error occurred while restoring auth info', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    command.action(logger, { options: {} }, () => {
      try {
        assert(commandCommandActionSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(auth.restoreAuth);
      }
    });
  });

  it('reuses current digestcontext when expireat is a future date', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    const futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: futureDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    command.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, false);

    try {
      assert(loggerLogSpy.notCalled);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('reuses current digestcontext when expireat is a future date (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    const futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: futureDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    command.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, true);

    try {
      assert(loggerLogSpy.notCalled);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('retrieves new digestcontext when no context present', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    command
      .ensureFormDigest('https://contoso.sharepoint.com', logger, undefined, false)
      .then(ctx => {
        try {
          assert.notStrictEqual(typeof ctx, 'undefined');
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => {
        done(e);
      });
  });

  it('retrieves updated digestcontext when expireat is past date', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    command.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, false);

    try {
      assert(loggerLogSpy.notCalled);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('retrieves updated digestcontext when expireat is past date (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    command.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, true);

    try {
      assert(loggerLogSpy.notCalled);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('handles error when contextinfo could not be retrieved (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.reject('Invalid request');
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    command.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, true).catch((err?: any) => {
      try {
        assert(err === "Invalid request");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves SPO URL from MS Graph when not retrieved previously', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return Promise.resolve({ webUrl: 'https://contoso.sharepoint.com' });
      }

      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    (command as any)
      .getSpoUrl(logger, false)
      .then((spoUrl: string) => {
        try {
          assert.strictEqual(spoUrl, 'https://contoso.sharepoint.com');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('retrieves SPO URL from MS Graph when not retrieved previously (debug)', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return Promise.resolve({ webUrl: 'https://contoso.sharepoint.com' });
      }

      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    (command as any)
      .getSpoUrl(logger, true)
      .then((spoUrl: string) => {
        try {
          assert.strictEqual(spoUrl, 'https://contoso.sharepoint.com');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns retrieved SPO URL when persisting connection info failed', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.reject());
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return Promise.resolve({ webUrl: 'https://contoso.sharepoint.com' });
      }

      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    (command as any)
      .getSpoUrl(logger, false)
      .then((spoUrl: string) => {
        try {
          assert.strictEqual(spoUrl, 'https://contoso.sharepoint.com');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns error when retrieving SPO URL failed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    (command as any)
      .getSpoUrl(logger, false)
      .then(() => {
        done('Expected error');
      }, (err: string) => {
        try {
          assert.strictEqual(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns error when retrieving SPO admin URL failed', (done) => {
    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    sinon.stub(command as any, 'getSpoUrl').callsFake(() => Promise.reject('An error has occurred'));

    (command as any)
      .getSpoAdminUrl(logger, false)
      .then(() => {
        done('Expected error');
      }, (err: string) => {
        try {
          assert.strictEqual(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('retrieves tenant ID when not retrieved previously', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return Promise.resolve(JSON.stringify([{
          _ObjectIdentity_: 'tenantId'
        }]));
      }

      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    sinon.stub(command as any, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'abc' }));

    (command as any)
      .getTenantId(logger, false)
      .then((tenantId: string) => {
        try {
          assert.strictEqual(tenantId, 'tenantId');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('retrieves tenant ID when not retrieved previously (debug)', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return Promise.resolve(JSON.stringify([{
          _ObjectIdentity_: 'tenantId'
        }]));
      }

      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    sinon.stub(command as any, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'abc' }));

    (command as any)
      .getTenantId(logger, true)
      .then((tenantId: string) => {
        try {
          assert.strictEqual(tenantId, 'tenantId');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns retrieved tenant ID when persisting connection info failed', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return Promise.resolve(JSON.stringify([{
          _ObjectIdentity_: 'tenantId'
        }]));
      }

      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    sinon.stub(command as any, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'abc' }));

    (command as any)
      .getTenantId(logger, false)
      .then((tenantId: string) => {
        try {
          assert.strictEqual(tenantId, 'tenantId');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns error when retrieving tenant ID failed', (done) => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    sinon.stub(command as any, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'abc' }));

    (command as any)
      .getTenantId(logger, false)
      .then(() => {
        done('Error expected');
      }, (err: any) => {
        try {
          assert.strictEqual(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('passes validation of unknown properties when no unknown properties are set', () => {
    const command = new MockCommand();
    assert.strictEqual(command.validateUnknownOptionsPublic({}, 'web', 'set'), true);
  });

  it('passes validation of unknown properties when valid unknown properties specified', () => {
    const command = new MockCommand();
    assert.strictEqual(command.validateUnknownOptionsPublic({ AllowAutomaticASPXPageIndexing: true }, 'web', 'set'), true);
  });

  it('fails validation of unknown properties when invalid unknown property specified', () => {
    const command = new MockCommand();
    assert.notStrictEqual(command.validateUnknownOptionsPublic({ AllowCreateDeclarativeWorkflow: true }, 'web', 'set'), true);
  });

  it('fails validation of unknown properties when unknown property of unsupported type specified', () => {
    const command = new MockCommand();
    assert.notStrictEqual(command.validateUnknownOptionsPublic({ AssociatedMemberGroup: {} }, 'web', 'set'), true);
  });

  it('returns default list of names of options with URLs if no names to exclude defined', () => {
    const expected = [
      'appCatalogUrl',
      'siteUrl',
      'webUrl',
      'origin',
      'url',
      'imageUrl',
      'actionUrl',
      'logoUrl',
      'libraryUrl',
      'thumbnailUrl',
      'targetUrl',
      'newSiteUrl',
      'previewImageUrl',
      'NoAccessRedirectUrl',
      'StartASiteFormUrl',
      'OrgNewsSiteUrl',
      'parentWebUrl',
      'siteLogoUrl'
    ];
    const command = new MockCommand();
    const actual = command.getNamesOfOptionsWithUrlsPublic();
    assert.deepStrictEqual(actual, expected);
  });

  it('returns filtered list of names of options with URLs when names to exclude defined', () => {
    const expected = [
      'appCatalogUrl',
      'siteUrl',
      'webUrl',
      'origin',
      'imageUrl',
      'actionUrl',
      'logoUrl',
      'libraryUrl',
      'thumbnailUrl',
      'targetUrl',
      'newSiteUrl',
      'previewImageUrl',
      'NoAccessRedirectUrl',
      'StartASiteFormUrl',
      'OrgNewsSiteUrl',
      'parentWebUrl',
      'siteLogoUrl'
    ];
    const command = new MockCommand();
    sinon.stub(command as any, 'getExcludedOptionsWithUrls').callsFake(() => ['url']);
    const actual = command.getNamesOfOptionsWithUrlsPublic();
    assert.deepStrictEqual(actual, expected);
  });

  it('resolves server-relative URLs in known options to absolute when SPO URL available', (done) => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      url: '/'
    };
    command
      .processOptions(options)
      .then(() => {
        try {
          assert.strictEqual(options.url, 'https://contoso.sharepoint.com/');
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('leaves absolute URLs as-is', (done) => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      url: 'https://contoso.sharepoint.com/sites/contoso'
    };
    command
      .processOptions(options)
      .then(() => {
        try {
          assert.strictEqual(options.url, 'https://contoso.sharepoint.com/sites/contoso');
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('leaves site-relative URLs as-is', (done) => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      url: 'sites/contoso'
    };
    command
      .processOptions(options)
      .then(() => {
        try {
          assert.strictEqual(options.url, 'sites/contoso');
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('leaves server-relative URLs as-is in unknown options', (done) => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      nonProcessedUrl: '/'
    };
    command
      .processOptions(options)
      .then(() => {
        try {
          assert.strictEqual(options.nonProcessedUrl, '/');
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('throws error when server-relative URL specified but SPO URL not available', (done) => {
    const command = new MockCommand();
    const options = {
      url: '/'
    };
    command
      .processOptions(options)
      .then(_ => {
        done('Options resolved while error expected');
      }, _ => done());
  });

  it('Shows an error when CLI is connected with authType "Secret"', (done) => {
    sinon.stub(auth.service, 'authType').value(5);

    const mock = new MockCommand();
    mock.action(logger, { options: {} }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('SharePoint does not support authentication using client ID and secret. Please use a different login type to use SharePoint commands.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});