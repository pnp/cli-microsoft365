import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import auth from '../../Auth';
import { Cli } from '../../cli/Cli';
import { CommandInfo } from '../../cli/CommandInfo';
import { Logger } from '../../cli/Logger';
import { CommandError } from '../../Command';
import request from '../../request';
import { telemetry } from '../../telemetry';
import { pid } from '../../utils/pid';
import { session } from '../../utils/session';
import { sinonUtil } from '../../utils/sinonUtil';
import DateAndPeriodBasedReport from './DateAndPeriodBasedReport';

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

describe('PeriodBasedReport', () => {
  const mockCommand = new MockCommand();
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(mockCommand);
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
    sinonUtil.restore([
      request.get,
      fs.writeFileSync
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(mockCommand.name.startsWith('mock'), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(mockCommand.description, null);
  });

  it('fails validation if period option is not passed', async () => {
    const actual = mockCommand.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation on invalid period', async () => {
    const actual = mockCommand.validate({ options: { period: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation on invalid date', async () => {
    const actual = mockCommand.validate({ options: { date: '10.10.2019' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation on valid \'D7\' period', async () => {
    const actual = await mockCommand.validate({
      options: {
        period: 'D7'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'D30\' period', async () => {
    const actual = await mockCommand.validate({
      options: {
        period: 'D30'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'D90\' period', async () => {
    const actual = await mockCommand.validate({
      options: {
        period: 'D90'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'180\' period', async () => {
    const actual = await mockCommand.validate({
      options: {
        period: 'D90'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if both period and date options set', async () => {
    const actual = await mockCommand.validate({ options: { period: 'D7', date: '2019-07-13' } }, commandInfo);
    assert.notStrictEqual(actual, true);
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

    await mockCommand.action(logger, { options: { period: 'D7' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('fails validation if the date option is not a valid date string', async () => {
    const actual = await mockCommand.validate({
      options:
      {
        date: '2018-X-09'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
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

    await mockCommand.action(logger, { options: { date: '2019-07-13' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(date=2019-07-13)");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(mockCommand.action(logger, { options: { period: 'D7' } } as any),
      new CommandError('An error has occurred'));
  });
});
