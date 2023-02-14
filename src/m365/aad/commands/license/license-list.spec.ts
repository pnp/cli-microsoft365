import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./license-list');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      request.get
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
    sinon.stub(request, 'get').callsFake(async () => { throw error; });

    await assert.rejects(command.action(logger, {
      options: {}
    }), new CommandError(error.error.message));
  });
});
