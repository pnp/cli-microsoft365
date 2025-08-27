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
import command from './report-pstncalls.js';

describe(commands.REPORT_PSTNCALLS, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const jsonOutput = {
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
    }]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    assert.strictEqual(command.name, commands.REPORT_PSTNCALLS);
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

  it('gets pstncalls in teams', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/communications/callRecords/getPstnCalls(fromDateTime=2019-11-01,toDateTime=2019-12-01)`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { fromDateTime: '2019-11-01', toDateTime: '2019-12-01' } });
    assert(loggerLogSpy.calledWith(jsonOutput.value));
  });

  it('gets pstncalls in teams with no toDateTime specified', async () => {
    const now = new Date();
    const fakeTimers = sinon.useFakeTimers(now);
    const toDateTime: string = formatting.encodeQueryParameter(now.toISOString());

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/communications/callRecords/getPstnCalls(fromDateTime=2019-11-01,toDateTime=${toDateTime})`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { fromDateTime: '2019-11-01' } });
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

    await assert.rejects(command.action(logger, { options: { fromDateTime: '2019-11-01', toDateTime: '2019-12-01' } } as any), new CommandError('An error has occurred'));
  });
});
