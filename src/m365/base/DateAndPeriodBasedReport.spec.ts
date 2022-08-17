import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
import { Cli, CommandInfo, Logger } from '../../cli';
import { CommandError } from '../../Command';
import request from '../../request';
import { sinonUtil } from '../../utils';
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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    sinonUtil.restore([
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

  it('get unique device type in teams and export it in a period', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    mockCommand.action(logger, { options: { debug: false, period: 'D7' } }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
        assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('gets details about Microsoft Teams user activity by user for the given date', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(date=2019-07-13)`) {
        return Promise.resolve(`
        Report Refresh Date,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Assigned Products,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Has Other Action,Report Period
        2019-08-14,abisha@contoso.onmicrosoft.com,,False,,,0,0,0,0,No,7
        2019-08-14,same@contoso.onmicrosoft.com,2019-05-22,False,,OFFICE 365 E3 DEVELOPER+MICROSOFT FLOW FREE,0,0,0,0,No,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    mockCommand.action(logger, { options: { debug: false, date: '2019-07-13' } }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(date=2019-07-13)");
        assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    mockCommand.action(logger, { options: { debug: false, period: 'D7' } } as any, (err?: any) => {
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
    const options = mockCommand.options;
    let containsOption = false;
    options.forEach((o: any) => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});