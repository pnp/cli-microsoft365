import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { misc } from '../../../../utils/misc.js';
import { MockRequests } from '../../../../utils/MockRequest.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './business-list.js';

export const mocks = {
  businesses: {
    request: {
      url: `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`
    },
    response: {
      body: {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#solutions/bookingBusinesses",
        "value": [
          {
            "id": "FourthCoffee@contoso.onmicrosoft.com",
            "displayName": "Fourth Coffee"
          }
        ]
      }
    }
  }
} satisfies MockRequests;

describe(commands.BUSINESS_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    assert.strictEqual(command.name, commands.BUSINESS_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName']);
  });

  it('lists Microsoft Bookings businesses in the tenant (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.businesses.request.url) {
        return misc.deepClone(mocks.businesses.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(
      [
        {
          "id": "FourthCoffee@contoso.onmicrosoft.com",
          "displayName": "Fourth Coffee"
        }
      ]
    ));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
