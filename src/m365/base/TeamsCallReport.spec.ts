import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
import { Logger } from '../../cli';
import { CommandError } from '../../Command';
import request from '../../request';
import Utils from '../../Utils';
import TeamsCallReport from './TeamsCallReport';

class MockCommand extends TeamsCallReport {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public get usageEndpoint(): string {
    return 'MockEndPoint';
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }
}

describe('TeamsCallReport', () => {
  const mockCommand = new MockCommand();
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
    (mockCommand as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      fs.writeFileSync
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
    assert.strictEqual(mockCommand.name.startsWith('mock'), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(mockCommand.description, null);
  });

  it('fails validation on invalid fromDateTime', () => {
    const actual = mockCommand.validate({
      options: {
        fromDateTime: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation on invalid toDateTime', () => {
    const actual = mockCommand.validate({
      options: {
        fromDateTime: '2020-12-01',
        toDateTime: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation on number of days between fromDateTime and toDateTme exceeding 90', () => {
    const actual = mockCommand.validate({
      options: {
        fromDateTime: '2020-08-01',
        toDateTime: '2020-12-01'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation on valid fromDateTime', () => {
    const actual = mockCommand.validate({
      options: {
        fromDateTime: '2020-12-01'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid fromDateTime and toDateTime', () => {
    const actual = mockCommand.validate({
      options: {
        fromDateTime: '2020-11-01',
        toDateTime: '2020-12-01'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('get pstncalls in teams', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/communications/callRecords/MockEndPoint(fromDateTime=2019-11-01,toDateTime=2019-12-01)`) {
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

    mockCommand.action(logger, { options: { debug: false, fromDateTime: '2019-11-01', toDateTime: '2019-12-01' } }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/beta/communications/callRecords/MockEndPoint(fromDateTime=2019-11-01,toDateTime=2019-12-01)");
        assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('get pstncalls in teams output in txt', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/communications/callRecords/MockEndPoint(fromDateTime=2019-11-01,toDateTime=2019-12-01)`) {
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

    mockCommand.action(logger, { options: { debug: false, fromDateTime: '2019-11-01', toDateTime: '2019-12-01' } }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/beta/communications/callRecords/MockEndPoint(fromDateTime=2019-11-01,toDateTime=2019-12-01)");
        assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('get pstncalls in teams output in json', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/communications/callRecords/MockEndPoint(fromDateTime=2019-11-01,toDateTime=2019-12-01)`) {
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

    mockCommand.action(logger, { options: { debug: false, fromDateTime: '2019-11-01', toDateTime: '2019-12-01', output: 'json' } }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/beta/communications/callRecords/MockEndPoint(fromDateTime=2019-11-01,toDateTime=2019-12-01)");
        assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => Promise.reject('An error has occurred'));

    mockCommand.action(logger, { options: { debug: false, fromDateTime: '2019-11-01', toDateTime: '2019-12-01' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = mockCommand.options();
    let containsOption = false;
    options.forEach((o: any) => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});