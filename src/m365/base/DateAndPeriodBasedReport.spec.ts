import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { telemetry } from '../../telemetry.js';
import auth from '../../Auth.js';
import { cli } from '../../cli/cli.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { Logger } from '../../cli/Logger.js';
import { CommandError } from '../../Command.js';
import request from '../../request.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import DateAndPeriodBasedReport, { dateAndPeriodBasedReportOptions } from './DateAndPeriodBasedReport.js';

class MockCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public get usageEndpoint(): string {
    return 'MockEndPoint';
  }

  public commandHelp(): void {
  }
}

describe('DateAndPeriodBasedReport', () => {
  const mockCommand = new MockCommand();
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof dateAndPeriodBasedReportOptions;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(mockCommand);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof dateAndPeriodBasedReportOptions;
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
    (mockCommand as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      fs.writeFileSync
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(mockCommand.name.startsWith('mock'), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(mockCommand.description, null);
  });

  it('fails validation if neither period nor date option is passed', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation on invalid period', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation on invalid date', () => {
    const actual = commandOptionsSchema.safeParse({ date: '10.10.2019' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation on valid \'D7\' period', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D7' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation on valid \'D30\' period', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D30' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation on valid \'D90\' period', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D90' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation on valid \'D180\' period', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D180' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both period and date options set', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D7', date: '2019-07-13' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D7', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('get unique device type in teams and export it in a period', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return `
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `;
      }

      throw 'Invalid request';
    });

    await mockCommand.action(logger, { options: commandOptionsSchema.parse({ period: 'D7' }) });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('fails validation if the date option is not a valid date string', () => {
    const actual = commandOptionsSchema.safeParse({ date: '2018-X-09' });
    assert.strictEqual(actual.success, false);
  });

  it('gets details about Microsoft Teams user activity by user for the given date', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(date=2019-07-13)`) {
        return `
        Report Refresh Date,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Assigned Products,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Has Other Action,Report Period
        2019-08-14,abisha@contoso.onmicrosoft.com,,False,,,0,0,0,0,No,7
        2019-08-14,same@contoso.onmicrosoft.com,2019-05-22,False,,OFFICE 365 E3 DEVELOPER+MICROSOFT FLOW FREE,0,0,0,0,No,7
        `;
      }

      throw 'Invalid request';
    });

    await mockCommand.action(logger, { options: commandOptionsSchema.parse({ date: '2019-07-13' }) });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(date=2019-07-13)");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(mockCommand.action(logger, { options: commandOptionsSchema.parse({ period: 'D7' }) }),
      new CommandError('An error has occurred'));
  });
});
