import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./oauth2grant-remove');

describe(commands.OAUTH2GRANT_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    promptOptions = undefined;
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.OAUTH2GRANT_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes OAuth2 permission grant when prompt confirmed (debug)', (done) => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants/YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, { options: { debug: true, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes OAuth2 permission grant when prompt confirmed', (done) => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants/YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, { options: { debug: false, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes OAuth2 permission grant when confirm specified', (done) => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants/YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek', confirm: true } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing OAuth2 permission grant when confirm option not passed', (done) => {
    command.action(logger, { options: { debug: false, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing OAuth2 permission grant when confirm option not passed (debug)', (done) => {
    command.action(logger, { options: { debug: true, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing OAuth2 permission grant when prompt not confirmed', (done) => {
    const deleteSpy = sinon.spy(request, 'delete');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });

    command.action(logger, { options: { debug: false, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } }, () => {
      try {
        assert(deleteSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing OAuth2 permission grant when prompt not confirmed (debug)', (done) => {
    const deleteSpy = sinon.spy(request, 'delete');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });

    command.action(logger, { options: { debug: true, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } }, () => {
      try {
        assert(deleteSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, confirm: true, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying grantId', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--grantId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying confirmation flag', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--confirm') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});