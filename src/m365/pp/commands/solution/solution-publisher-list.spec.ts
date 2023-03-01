import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { powerPlatform } from '../../../../utils/powerPlatform';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./solution-publisher-list');

describe(commands.SOLUTION_PUBLISHER_LIST, () => {
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const publisherResponse = {
    "value": [
      {
        "publisherid": "00000001-0000-0000-0000-00000000005a",
        "uniquename": "Cree38e",
        "friendlyname": "CDS Default Publisher",
        "versionnumber": 1074060,
        "isreadonly": false,
        "description": null,
        "customizationprefix": "cr6c3",
        "customizationoptionvalueprefix": 43186
      },
      {
        "publisherid": "d21aab70-79e7-11dd-8874-00188b01e34f",
        "uniquename": "MicrosoftCorporation",
        "friendlyname": "MicrosoftCorporation",
        "versionnumber": 1226559,
        "isreadonly": false,
        "customizationprefix": "",
        "customizationoptionvalueprefix": 0
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SOLUTION_PUBLISHER_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['publisherid', 'uniquename', 'friendlyname']);
  });

  it('retrieves publishers from power platform environment', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&$filter=publisherid ne 'd21aab70-79e7-11dd-8874-00188b01e34f'&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
          return publisherResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: validEnvironment } });
    assert(loggerLogSpy.calledWith(publisherResponse.value));
  });

  it('retrieves publishers from power platform environment including the Microsoft Publishers', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
          return publisherResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: validEnvironment, includeMicrosoftPublishers: true } });
    assert(loggerLogSpy.calledWith(publisherResponse.value));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&$filter=publisherid ne 'd21aab70-79e7-11dd-8874-00188b01e34f'&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
          throw {
            error: {
              'odata.error': {
                code: '-1, InvalidOperationException',
                message: {
                  value: `Resource '' does not exist or one of its queried reference-property objects are not present`
                }
              }
            }
          };
        }
      }

    });

    await assert.rejects(command.action(logger, { options: { environment: validEnvironment } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
