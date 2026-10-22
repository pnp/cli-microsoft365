import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './hidedefaultthemes-get.js';

describe(commands.HIDEDEFAULTTHEMES_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
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
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
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
      options: commandOptionsSchema.parse({})
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
      options: commandOptionsSchema.parse({
        debug: true
      })
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

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, verbose: true }) });
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, verbose: true }) }),
      new CommandError(error.error['odata.error'].message.value));
  });

  it('passes validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });
});
