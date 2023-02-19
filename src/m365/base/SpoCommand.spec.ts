import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../telemetry';
import auth from '../../Auth';
import { Logger } from '../../cli/Logger';
import { CommandError } from '../../Command';
import request from '../../request';
import { pid } from '../../utils/pid';
import { sinonUtil } from '../../utils/sinonUtil';
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

  public async commandAction(): Promise<void> {
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
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      auth.storeConnectionInfo,
      auth.restoreAuth
    ]);
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('correctly reports an error while restoring auth info', async () => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();

    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('doesn\'t execute command when error occurred while restoring auth info', async () => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await assert.rejects(command.action(logger, { options: {} }));
    assert(commandCommandActionSpy.notCalled);
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
      'actionUrl',
      'imageUrl',
      'libraryUrl',
      'logoUrl',
      'newSiteUrl',
      'NoAccessRedirectUrl',
      'OrgNewsSiteUrl',
      'origin',
      'parentUrl',
      'parentWebUrl',
      'previewImageUrl',
      'siteLogoUrl',
      'siteUrl',
      'StartASiteFormUrl',
      'targetUrl',
      'thumbnailUrl',
      'url',
      'webUrl'
    ];
    const command = new MockCommand();
    const actual = command.getNamesOfOptionsWithUrlsPublic();
    assert.deepStrictEqual(actual, expected);
  });

  it('returns filtered list of names of options with URLs when names to exclude defined', () => {
    const expected = [
      'appCatalogUrl',
      'actionUrl',
      'imageUrl',
      'libraryUrl',
      'logoUrl',
      'newSiteUrl',
      'NoAccessRedirectUrl',
      'OrgNewsSiteUrl',
      'origin',
      'parentUrl',
      'parentWebUrl',
      'previewImageUrl',
      'siteLogoUrl',
      'siteUrl',
      'StartASiteFormUrl',
      'targetUrl',
      'thumbnailUrl',
      'webUrl'
    ];
    const command = new MockCommand();
    sinon.stub(command as any, 'getExcludedOptionsWithUrls').callsFake(() => ['url']);
    const actual = command.getNamesOfOptionsWithUrlsPublic();
    assert.deepStrictEqual(actual, expected);
  });

  it('resolves server-relative URLs in known options to absolute when SPO URL available', async () => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      url: '/'
    };
    await command.processOptions(options);
    assert.strictEqual(options.url, 'https://contoso.sharepoint.com/');
  });

  it('leaves absolute URLs as-is', async () => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      url: 'https://contoso.sharepoint.com/sites/contoso'
    };
    await command.processOptions(options);
    assert.strictEqual(options.url, 'https://contoso.sharepoint.com/sites/contoso');
  });

  it('leaves site-relative URLs as-is', async () => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      url: 'sites/contoso'
    };
    await command.processOptions(options);
    assert.strictEqual(options.url, 'sites/contoso');
  });

  it('leaves server-relative URLs as-is in unknown options', async () => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      nonProcessedUrl: '/'
    };
    await command.processOptions(options);
    assert.strictEqual(options.nonProcessedUrl, '/');
  });

  it('throws error when server-relative URL specified but SPO URL not available', async () => {
    const command = new MockCommand();
    const options = {
      url: '/'
    };
    await assert.rejects(command.processOptions(options));
  });

  it('Shows an error when CLI is connected with authType "Secret"', async () => {
    sinon.stub(auth.service, 'authType').value(5);

    const mock = new MockCommand();
    await assert.rejects(mock.action(logger, { options: {} }),
      new CommandError('SharePoint does not support authentication using client ID and secret. Please use a different login type to use SharePoint commands.'));
  });
});
