import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './report-directroutingcalls.js';

describe(commands.REPORT_DIRECTROUTINGCALLS, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

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
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    commandInfo = Cli.getCommandInfo(command);
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.REPORT_DIRECTROUTINGCALLS);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'calleeNumber', 'callerNumber', 'startDateTime']);
  });

  it('fails validation on invalid fromDateTime', async () => {
    const actual = await command.validate({
      options: {
        fromDateTime: 'abc'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation on invalid toDateTime', async () => {
    const actual = await command.validate({
      options: {
        fromDateTime: '2020-12-01',
        toDateTime: 'abc'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation on number of days between fromDateTime and toDateTme exceeding 90', async () => {
    const actual = await command.validate({
      options: {
        fromDateTime: '2020-08-01',
        toDateTime: '2020-12-01'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation on valid fromDateTime', async () => {
    const validfromDateTime: any = new Date();
    //fromDateTime should be less than 90 days ago for passing validation
    validfromDateTime.setDate(validfromDateTime.getDate() - 70);
    const actual = await command.validate({
      options: {
        fromDateTime: validfromDateTime.toISOString().substr(0, 10)
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid fromDateTime and toDateTime', async () => {
    const actual = await command.validate({
      options: {
        fromDateTime: '2020-11-01',
        toDateTime: '2020-12-01'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('gets directroutingcalls in teams', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=2019-11-01,toDateTime=2019-12-01)`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { fromDateTime: '2019-11-01', toDateTime: '2019-12-01' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=2019-11-01,toDateTime=2019-12-01)");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('gets directroutingcalls in teams with no toDateTime specified', async () => {
    const now = new Date();
    const fakeTimers = sinon.useFakeTimers(now);
    const toDateTime: string = formatting.encodeQueryParameter(now.toISOString());

    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=2019-11-01,toDateTime=${toDateTime})`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { fromDateTime: '2019-11-01' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, `https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=2019-11-01,toDateTime=${toDateTime})`);
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
    fakeTimers.restore();
  });

  it('correctly handles random API error', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };

    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, { options: { fromDateTime: '2019-11-01', toDateTime: '2019-12-01' } } as any), new CommandError('An error has occurred'));
  });
});
