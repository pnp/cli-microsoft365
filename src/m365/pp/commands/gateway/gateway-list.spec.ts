import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./gateway-list');

describe(commands.GATEWAY_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GATEWAY_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name']);
  });

  it('retrieves list of gateways (debug)', (done) => {
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
          return Promise.resolve(gateways);
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith(gateways.value));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      });
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
