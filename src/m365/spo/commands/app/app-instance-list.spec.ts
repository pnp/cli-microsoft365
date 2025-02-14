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
import command from './app-instance-list.js';

describe(commands.APP_INSTANCE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_INSTANCE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), [`Title`, `AppId`]);
  });

  it('fails validation when siteUrl is not a valid url', async () => {
    const actual = await command.validate({ options: { siteUrl: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the siteUrl is a valid url', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/testsite' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves available apps from the site collection', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/AppTiles') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            value: [
              {
                AppId: 'b2307a39-e878-458b-bc90-03bc578531d6',
                Title: 'online-client-side-solution'
              },
              {
                AppId: 'e5f65aef-68fe-45b0-801e-92733dd57e2c',
                Title: 'onprem-client-side-solution'
              }
            ]
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/testsite' } });
    assert(loggerLogSpy.calledWith([
      {
        AppId: 'b2307a39-e878-458b-bc90-03bc578531d6',
        Title: 'online-client-side-solution'
      },
      {
        AppId: 'e5f65aef-68fe-45b0-801e-92733dd57e2c',
        Title: 'onprem-client-side-solution'
      }
    ]));
  });



  it('correctly handles no apps found in the site collection', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/AppTiles') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { value: [] };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/testsite', verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly([]));
  });

  it('correctly handles error while listing apps in the site collection', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/testsite'
      }
    } as any), new CommandError('An error has occurred'));
  });
});
