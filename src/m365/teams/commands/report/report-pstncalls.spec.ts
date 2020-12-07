import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./report-pstncalls');

describe(commands.TEAMS_REPORT_PSTNCALLS, () => {
  let log: string[];
  let logger: Logger;

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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TEAMS_REPORT_PSTNCALLS), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('get details about PSTN calls made within a given time period', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/communications/callRecords/getPstnCalls(fromDateTime=2019-11-01,toDateTime=2019-12-01)`) {
        return Promise.resolve(
          {
            "@odata.context": "https://graph.microsoft.com/beta/$metadata#Collection(microsoft.graph.callRecords.pstnCallLogRow)",
            "@odata.count": 1000,
            "value": [{
              "id": "9c4984c7-6c3c-427d-a30c-bd0b2eacee90",
              "callId": "1835317186_112562680@61.221.3.176",
              "userId": "db03c14b-06eb-4189-939b-7cbf3a20ba27",
              "userPrincipalName": "richard.malk@contoso.com",
              "userDisplayName": "Richard Malk",
              "startDateTime": "2019-11-01T00:00:08.2589935Z",
              "endDateTime": "2019-11-01T00:03:47.2589935Z",
              "duration": 219,
              "charge": 0.00,
              "callType": "user_in",
              "currency": "USD",
              "calleeNumber": "+1234567890",
              "usageCountryCode": "US",
              "tenantCountryCode": "US",
              "connectionCharge": 0.00,
              "callerNumber": "+0123456789",
              "destinationContext": null,
              "destinationName": "United States",
              "conferenceId": null,
              "licenseCapability": "MCOPSTNU",
              "inventoryType": "Subscriber"
            }],
            "@odata.nextLink": "https://graph.microsoft.com/beta/communications/callRecords/getPstnCalls(from=2019-11-01,to=2019-12-01)?$skip=1000"
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, fromDateTime: '2019-11-01', toDateTime: '2019-12-01' } }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/beta/communications/callRecords/getPstnCalls(fromDateTime=2019-11-01,toDateTime=2019-12-01)");
        assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});