import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './gateway-list.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.GATEWAY_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
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
    assert.strictEqual(command.name, commands.GATEWAY_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name']);
  });

  it('retrieves list of gateways (debug)', async () => {
    const gateways: any = {
      "value": [
        {
          "id": "22660b34-31b3-4744-a99c-5e154458a784",
          "gatewayId": 0,
          "name": "Contoso Gateway",
          "type": "Resource",
          "publicKey": {
            "exponent": "AQAB",
            "modulus": "okJBN8MJyaVkjfkN75B6OgP7RYiC3KFMFaky9KqqudqiTOcZPRXlsG+emrbnnBpFzw7ywe4gWtUGnPCqy01RKeDZrFA3QfkVPJpH28OWfrmgkMQNsI4Op2uxwEyjnJAyfYxIsHlpevOZoDKpWJgV+sH6MRf/+LK4hN3vNJuWKKpf90rNwjipnYMumHyKVkd4Vssc9Ftsu4Samu0/TkXzUkyje5DxMF2ZK1Nt2TgItBcpKi4wLCP4bPDYYaa9vfOmBlji7U+gwuE5bjnmjazFljQ5sOP0VdA0fRoId3+nI7n1rSgRq265jNHX84HZbm2D/Pk8C0dElTmYEswGPDWEJQ=="
          },
          "gatewayAnnotation": "{\"gatewayContactInformation\":[\"admin@contoso.onmicrosoft.com\"],\"gatewayVersion\":\"3000.122.8\",\"gatewayWitnessString\":\"{\\\"EncryptedResult\\\":\\\"UyfEqNSy0e9S4D0m9oacPyYhgiXLWusCiKepoLudnTEe68iw9qEaV6qNqTbSKlVUwUkD9KjbnbV0O3vU97Q/KTJXpw9/1SiyhpO+JN1rcaL51mPjyQo0WwMHMo2PU3rdEyxsLjkJxJZHTh4+XGB/lQ==\\\",\\\"IV\\\":\\\"QxCYjHEl8Ab9i78ZBYpnDw==\\\",\\\"Signature\\\":\\\"upVXK3DvWdj5scw8iUDDilzQz1ovuNgeuXRpmf0N828=\\\"}\",\"gatewayMachine\":\"SPFxDevelop\",\"gatewaySalt\":\"rA1M34AdgdCbOYQMvo/izA==\",\"gatewayWitnessStringLegacy\":null,\"gatewaySaltLegacy\":null,\"gatewayDepartment\":null,\"gatewayVirtualNetworkSubnetId\":null}"
        }
      ]
    };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/myorg/gateways`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json;odata.metadata=none') === 0) {
          return gateways;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(gateways.value));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
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
    });

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
