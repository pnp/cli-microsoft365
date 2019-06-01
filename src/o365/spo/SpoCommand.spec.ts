import * as sinon from 'sinon';
import * as assert from 'assert';
import SpoCommand from './SpoCommand';
import request from '../../request';
import auth, { Site } from './SpoAuth';
import Utils from '../../Utils';
import { CommandError } from '../../Command';
import { FormDigestInfo } from './spo';
import appInsights from '../../appInsights';

class MockCommand extends SpoCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }
}

describe('SpoCommand', () => {
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let cmdInstance: any;
  let log: string[];

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };

    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post,
    ]);
  });

  after(() => {
    Utils.restore([
      request.post,
      auth.ensureAccessToken,
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
  });

  it('correctly reports an error while restoring auth info', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();

    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    cmdInstance.action({ options: {} }, () => {
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
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    auth.site.tenantId = 'abc';

    let futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: futureDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    }

    command.ensureFormDigest(cmdInstance, ctx, false);

    try {
      assert(cmdInstanceLogSpy.notCalled);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('reuses current digestcontext when expireat is a future date (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    auth.site.tenantId = 'abc';

    let futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: futureDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    }

    command.ensureFormDigest(cmdInstance, ctx, true);

    try {
      assert(cmdInstanceLogSpy.notCalled);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('retrieves new digestcontext when no context present', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    auth.site.tenantId = 'abc';

    command
      .ensureFormDigest(cmdInstance, undefined, false)
      .then(ctx => {
        try {
          assert.notEqual(typeof ctx, 'undefined');
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
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    auth.site.tenantId = 'abc';

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    }

    command.ensureFormDigest(cmdInstance, ctx, false);

    try {
      assert(cmdInstanceLogSpy.notCalled);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('retrieves updated digestcontext when expireat is past date (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    auth.site.tenantId = 'abc';

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    }

    command.ensureFormDigest(cmdInstance, ctx, true);

    try {
      assert(cmdInstanceLogSpy.notCalled);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('handles error when contextinfo could not be retrieved (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.reject('Invalid request');
      }
      return Promise.reject('Invalid request');
    });

    const command = new MockCommand();
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    auth.site.tenantId = 'abc';

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    }

    command.ensureFormDigest(cmdInstance, ctx, true).catch((err?: any) => {
      try {
        assert(err === "Invalid request");
        done();
      }
      catch (e) {
        done(e);
      }
    })
  });
});