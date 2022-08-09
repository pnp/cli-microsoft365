import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
import { Logger } from '../../cli';
import { CommandError } from '../../Command';
import request from '../../request';
import { sinonUtil } from '../../utils';
import SpoCommand from './SpoCommand';

class MockCommand extends SpoCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  constructor() {
    super();

    this.options.unshift(
      {
        option: '--url [url]'
      },
      {
        option: '--nonProcessedUrl [nonProcessedUrl]'
      }
    );
  }

  public commandAction(): void {
  }

  public validateUnknownCsomOptionsPublic(options: any, csomObject: string, csomPropertyType: 'get' | 'set'): string | boolean {
    return this.validateUnknownCsomOptions(options, csomObject, csomPropertyType);
  }

  public getNamesOfOptionsWithUrlsPublic(): string[] {
    return this.getNamesOfOptionsWithUrls();
  }
}

describe('SpoCommand', () => {
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      auth.storeConnectionInfo
    ]);
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  after(() => {
    sinonUtil.restore([
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
        sinonUtil.restore(auth.restoreAuth);
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
        sinonUtil.restore(auth.restoreAuth);
      }
    });
  });

  it('passes validation of unknown properties when no unknown properties are set', async () => {
    const command = new MockCommand();
    assert.strictEqual(command.validateUnknownCsomOptionsPublic({}, 'web', 'set'), true);
  });

  it('passes validation of unknown properties when valid unknown properties specified', async () => {
    const command = new MockCommand();
    assert.strictEqual(command.validateUnknownCsomOptionsPublic({ AllowAutomaticASPXPageIndexing: true }, 'web', 'set'), true);
  });

  it('fails validation of unknown properties when invalid unknown property specified', async () => {
    const command = new MockCommand();
    assert.notStrictEqual(command.validateUnknownCsomOptionsPublic({ AllowCreateDeclarativeWorkflow: true }, 'web', 'set'), true);
  });

  it('fails validation of unknown properties when unknown property of unsupported type specified', async () => {
    const command = new MockCommand();
    assert.notStrictEqual(command.validateUnknownCsomOptionsPublic({ AssociatedMemberGroup: {} }, 'web', 'set'), true);
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