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
const command: Command = require('./homesite-get');

describe(commands.HOMESITE_GET, () => {
  let log: any[];
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HOMESITE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about the Home Site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/SP.SPHSite/Details') {
        return {
          "SiteId": "53ad95dc-5d2c-42a3-a63c-716f7b8014f5",
          "WebId": "288ce497-483c-4cd5-b8a2-27b726d002e2",
          "LogoUrl": "https://contoso.sharepoint.com/sites/Work/siteassets/work.png",
          "Title": "Work @ Contoso",
          "Url": "https://contoso.sharepoint.com/sites/Work"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} } as any);
    assert(loggerLogSpy.calledWith({
      "SiteId": "53ad95dc-5d2c-42a3-a63c-716f7b8014f5",
      "WebId": "288ce497-483c-4cd5-b8a2-27b726d002e2",
      "LogoUrl": "https://contoso.sharepoint.com/sites/Work/siteassets/work.png",
      "Title": "Work @ Contoso",
      "Url": "https://contoso.sharepoint.com/sites/Work"
    }));
  });

  it(`doesn't output anything when information about the Home Site is not available`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/SP.SPHSite/Details') {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles random API error', async () => {
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

    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(error.error['odata.error'].message.value));
  });
});
