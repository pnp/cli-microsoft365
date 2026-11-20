import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
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
import command, { options } from './report-directroutingcalls.js';

describe(commands.REPORT_DIRECTROUTINGCALLS, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const jsonOutput = {
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
    }]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    assert.strictEqual(command.name, commands.REPORT_DIRECTROUTINGCALLS);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'calleeNumber', 'callerNumber', 'startDateTime']);
  });

  it('fails validation on invalid fromDateTime', () => {
    const actual = commandOptionsSchema.safeParse({ fromDateTime: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation on invalid toDateTime', () => {
    const actual = commandOptionsSchema.safeParse({ fromDateTime: '2020-12-01', toDateTime: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation on number of days between fromDateTime and toDateTme exceeding 90', () => {
    const actual = commandOptionsSchema.safeParse({ fromDateTime: '2020-08-01', toDateTime: '2020-12-01' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation on valid fromDateTime', () => {
    const validfromDateTime: any = new Date();
    //fromDateTime should be less than 90 days ago for passing validation
    validfromDateTime.setDate(validfromDateTime.getDate() - 70);
    const actual = commandOptionsSchema.safeParse({ fromDateTime: validfromDateTime.toISOString().substr(0, 10) });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation on valid fromDateTime and toDateTime', () => {
    const actual = commandOptionsSchema.safeParse({ fromDateTime: '2020-11-01', toDateTime: '2020-12-01' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ fromDateTime: '2020-11-01', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('gets directroutingcalls in teams', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=2019-11-01,toDateTime=2019-12-01)`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ fromDateTime: '2019-11-01', toDateTime: '2019-12-01' }) });
    assert(loggerLogSpy.calledWith(jsonOutput.value));
  });

  it('gets directroutingcalls in teams with no toDateTime specified', async () => {
    const now = new Date();
    const fakeTimers = sinon.useFakeTimers(now);
    const toDateTime: string = formatting.encodeQueryParameter(now.toISOString());
    const fromDateTime: Date = new Date();
    fromDateTime.setDate(fromDateTime.getDate() - 70);
    const fromDateTimeString: string = fromDateTime.toISOString().substring(0, 10);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=${fromDateTimeString},toDateTime=${toDateTime})`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ fromDateTime: fromDateTimeString }) });
    assert(loggerLogSpy.calledWith(jsonOutput.value));
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ fromDateTime: '2019-11-01', toDateTime: '2019-12-01' }) }), new CommandError('An error has occurred'));
  });
});
