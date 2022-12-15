import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./oauth2grant-add');

describe(commands.OAUTH2GRANT_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.OAUTH2GRANT_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds OAuth2 permission grant (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants`) > -1) {
        if (opts.headers &&
          opts.headers['content-type'] &&
          (opts.headers['content-type'] as string).indexOf('application/json') === 0 &&
          opts.data.clientId === '6a7b1395-d313-4682-8ed4-65a6265a6320' &&
          opts.data.resourceId === '6a7b1395-d313-4682-8ed4-65a6265a6321' &&
          opts.data.scope === 'user_impersonation') {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6321', scope: 'user_impersonation' } } as any);
    assert(loggerLogToStderrSpy.called);
  });

  it('adds OAuth2 permission grant', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants`) > -1) {
        if (opts.headers &&
          opts.headers['content-type'] &&
          (opts.headers['content-type'] as string).indexOf('application/json') === 0 &&
          opts.data.clientId === '6a7b1395-d313-4682-8ed4-65a6265a6320' &&
          opts.data.resourceId === '6a7b1395-d313-4682-8ed4-65a6265a6321' &&
          opts.data.scope === 'user_impersonation') {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6321', scope: 'user_impersonation' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
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

  it('fails validation if the clientId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { clientId: '123', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6320', scope: 'user_impersonation' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the resourceId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '123', scope: 'user_impersonation' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when clientId, resourceId and scope are specified', async () => {
    const actual = await command.validate({ options: { clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6320', scope: 'user_impersonation' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying clientId', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--clientId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying resourceId', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--resourceId') > -1) {
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
