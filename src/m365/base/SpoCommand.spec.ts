import * as sinon from 'sinon';
import * as assert from 'assert';
import SpoCommand from './SpoCommand';
import request from '../../request';
import auth from '../../Auth';
import Utils from '../../Utils';
import { CommandError } from '../../Command';
import { FormDigestInfo } from '../spo/spo';
import appInsights from '../../appInsights';
import { CommandInstance } from '../../cli';

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

  public validateUnknownOptionsPublic(options: any, csomObject: string, csomPropertyType: 'get' | 'set'): string | boolean {
    return this.validateUnknownOptions(options, csomObject, csomPropertyType);
  }
}

describe('SpoCommand', () => {
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let cmdInstance: any;
  let log: string[];

  before(() => {
    auth.service.connected = true;
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  })

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
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
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

    let futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: futureDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    }

    command.ensureFormDigest('https://contoso.sharepoint.com', cmdInstance, ctx, false);

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
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
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

    let futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: futureDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    }

    command.ensureFormDigest('https://contoso.sharepoint.com', cmdInstance, ctx, true);

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
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
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

    command
      .ensureFormDigest('https://contoso.sharepoint.com', cmdInstance, undefined, false)
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
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    }

    command.ensureFormDigest('https://contoso.sharepoint.com', cmdInstance, ctx, false);

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
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
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

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    }

    command.ensureFormDigest('https://contoso.sharepoint.com', cmdInstance, ctx, true);

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
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
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

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    }

    command.ensureFormDigest('https://contoso.sharepoint.com', cmdInstance, ctx, true).catch((err?: any) => {
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
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    (command as any)
      .getSpoUrl(cmdInstance, false)
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
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    (command as any)
      .getSpoUrl(cmdInstance, true)
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
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    (command as any)
      .getSpoUrl(cmdInstance, false)
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
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };

    (command as any)
      .getSpoUrl(cmdInstance, false)
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
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };
    sinon.stub(command as any, 'getSpoUrl').callsFake(() => Promise.reject('An error has occurred'));

    (command as any)
      .getSpoAdminUrl(cmdInstance, false)
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
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };
    sinon.stub(command as any, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'abc' }));

    (command as any)
      .getTenantId(cmdInstance, false)
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
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };
    sinon.stub(command as any, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'abc' }));

    (command as any)
      .getTenantId(cmdInstance, true)
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
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };
    sinon.stub(command as any, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'abc' }));

    (command as any)
      .getTenantId(cmdInstance, false)
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
    sinon.stub(request, 'post').callsFake((opts) => Promise.reject('An error has occurred'));

    const command = new MockCommand();
    const cmdInstance = {
      commandWrapper: {
        command: 'spo command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };
    sinon.stub(command as any, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'abc' }));

    (command as any)
      .getTenantId(cmdInstance, false)
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
});