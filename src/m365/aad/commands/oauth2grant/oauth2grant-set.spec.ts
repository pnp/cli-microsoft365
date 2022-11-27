import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./oauth2grant-set');

describe(commands.OAUTH2GRANT_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.OAUTH2GRANT_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates OAuth2 permission grant (debug)', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants/YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek`) > -1) {
        if (opts.headers &&
          opts.headers['content-type'] &&
          (opts.headers['content-type'] as string).indexOf('application/json') === 0 &&
          opts.data.scope === 'user_impersonation') {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek', scope: 'user_impersonation' } });
    assert(loggerLogToStderrSpy.called);
  });

  it('updates OAuth2 permission grant', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants/YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek`) > -1) {
        if (opts.headers &&
          opts.headers['content-type'] &&
          (opts.headers['content-type'] as string).indexOf('application/json') === 0 &&
          opts.data.scope === 'user_impersonation') {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek', scope: 'user_impersonation' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'patch').callsFake(() => {
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

    await assert.rejects(command.action(logger, { options: { clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6320', scope: 'user_impersonation' } } as any),
      new CommandError('An error has occurred'));
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

  it('supports specifying scope', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--scope') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});