import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './license-list.js';

describe(commands.LICENSE_LIST, () => {
  //#region Mocked Responses
  const licenseResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscribedSkus",
    "value": [
      {
        "capabilityStatus": "Enabled",
        "consumedUnits": 14,
        "id": "48a80680-7326-48cd-9935-b556b81d3a4e_c7df2760-2c81-4ef7-b578-5b5392b571df",
        "prepaidUnits": {
          "enabled": 25,
          "suspended": 0,
          "warning": 0
        },
        "servicePlans": [
          {
            "servicePlanId": "8c098270-9dd4-4350-9b30-ba4703f3b36b",
            "servicePlanName": "ADALLOM_S_O365",
            "provisioningStatus": "Success",
            "appliesTo": "User"
          }
        ],
        "skuId": "c7df2760-2c81-4ef7-b578-5b5392b571df",
        "skuPartNumber": "ENTERPRISEPREMIUM",
        "appliesTo": "User"
      },
      {
        "capabilityStatus": "Suspended",
        "consumedUnits": 14,
        "id": "48a80680-7326-48cd-9935-b556b81d3a4e_d17b27af-3f49-4822-99f9-56a661538792",
        "prepaidUnits": {
          "enabled": 0,
          "suspended": 25,
          "warning": 0
        },
        "servicePlans": [
          {
            "servicePlanId": "f9646fb2-e3b2-4309-95de-dc4833737456",
            "servicePlanName": "CRMSTANDARD",
            "provisioningStatus": "Disabled",
            "appliesTo": "User"
          }
        ],
        "skuId": "d17b27af-3f49-4822-99f9-56a661538792",
        "skuPartNumber": "CRMSTANDARD",
        "appliesTo": "User"
      }
    ]
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LICENSE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'skuId', 'skuPartNumber']);
  });

  it('retrieves licenses', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/subscribedSkus`)) {
        return licenseResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(licenseResponse.value));

  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The licenses cannot be found.'
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {}
    }), new CommandError(error.error.message));
  });
});
