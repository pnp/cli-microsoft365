import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./report-directroutingcalls');

describe(commands.REPORT_DIRECTROUTINGCALLS, () => {
  let log: string[];
  let logger: Logger;

  const jsonOutput = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.callRecords.directRoutingLogRow)",
    "@odata.count": 1000,
    "value": [{
      "id": "9e8bba57-dc14-533a-a7dd-f0da6575eed1",
      "correlationId": "c98e1515-a937-4b81-b8a8-3992afde64e0",
      "userId": "db03c14b-06eb-4189-939b-7cbf3a20ba27",
      "userPrincipalName": "richard.malk@contoso.com",
      "userDisplayName": "Richard Malk",
      "startDateTime": "2019-11-01T00:00:25.105Z",
      "inviteDateTime": "2019-11-01T00:00:21.949Z",
      "failureDateTime": "0001-01-01T00:00:00Z",
      "endDateTime": "2019-11-01T00:00:30.105Z",
      "duration": 5,
      "callType": "ByotIn",
      "successfulCall": true,
      "callerNumber": "+12345678***",
      "calleeNumber": "+01234567***",
      "mediaPathLocation": "USWE",
      "signalingLocation": "EUNO",
      "finalSipCode": 0,
      "callEndSubReason": 540000,
      "finalSipCodePhrase": "BYE",
      "trunkFullyQualifiedDomainName": "tll-audiocodes01.adatum.biz",
      "mediaBypassEnabled": false
    }],
    "@odata.nextLink": "https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=2019-11-01,toDateTime=2019-12-01)?$skip=1000"
  };

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
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.REPORT_DIRECTROUTINGCALLS), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'calleeNumber', 'callerNumber', 'startDateTime']);
  });

  it('fails validation on invalid fromDateTime', () => {
    const actual = command.validate({
      options: {
        fromDateTime: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation on invalid toDateTime', () => {
    const actual = command.validate({
      options: {
        fromDateTime: '2020-12-01',
        toDateTime: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation on number of days between fromDateTime and toDateTme exceeding 90', () => {
    const actual = command.validate({
      options: {
        fromDateTime: '2020-08-01',
        toDateTime: '2020-12-01'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation on valid fromDateTime', () => {
    const validfromDateTime: any = new Date();
    //fromDateTime should be less than 90 days ago for passing validation
    validfromDateTime.setDate(validfromDateTime.getDate() - 70);
    const actual = command.validate({
      options: {
        fromDateTime: validfromDateTime.toISOString().substr(0, 10)
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid fromDateTime and toDateTime', () => {
    const actual = command.validate({
      options: {
        fromDateTime: '2020-11-01',
        toDateTime: '2020-12-01'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('gets directroutingcalls in teams', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=2019-11-01,toDateTime=2019-12-01)`) {
        return Promise.resolve(jsonOutput);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, fromDateTime: '2019-11-01', toDateTime: '2019-12-01' } }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=2019-11-01,toDateTime=2019-12-01)");
        assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets directroutingcalls in teams with no toDateTime specified', (done) => {
    const now = new Date();
    const fakeTimers = sinon.useFakeTimers(now);
    const toDateTime: string = encodeURIComponent(now.toISOString());

    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=2019-11-01,toDateTime=${toDateTime})`) {
        return Promise.resolve(jsonOutput);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, fromDateTime: '2019-11-01' } }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, `https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=2019-11-01,toDateTime=${toDateTime})`);
        assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        fakeTimers.restore();
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: false, fromDateTime: '2019-11-01', toDateTime: '2019-12-01' } } as any, (err?: any) => {
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
    const options = command.options();
    let containsOption = false;
    options.forEach((o: any) => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});