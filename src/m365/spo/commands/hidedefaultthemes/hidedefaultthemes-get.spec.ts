import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./hidedefaultthemes-get');

describe(commands.HIDEDEFAULTTHEMES_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HIDEDEFAULTTHEMES_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('uses correct API url', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetHideDefaultThemes') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
      }
    });
  });

  it('uses correct API url (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetHideDefaultThemes') > -1) {
        return 'Correct Url';
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true
      }
    });
  });

  it('gets the current value of the HideDefaultThemes setting', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetHideDefaultThemes') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { value: true };
        }
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, verbose: true } });
    assert(loggerLogSpy.calledWith(true), 'Invalid request');
  });

  it('gets the current value of the HideDefaultThemes setting - handle error', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetHideDefaultThemes') > -1) {
        throw error;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, verbose: true } } as any),
      new CommandError(error.error['odata.error'].message.value));
  });
});
