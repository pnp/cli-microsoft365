import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { telemetry } from '../../telemetry';
import auth from '../../Auth';
import { Cli } from '../../cli/Cli';
import { CommandInfo } from '../../cli/CommandInfo';
import { Logger } from '../../cli/Logger';
import { CommandError } from '../../Command';
import request from '../../request';
import { pid } from '../../utils/pid';
import { sinonUtil } from '../../utils/sinonUtil';
import PeriodBasedReport from './PeriodBasedReport';

class MockCommand extends PeriodBasedReport {
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
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      telemetry.trackEvent,
      pid.getProcessName,
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

  it('fails validation on invalid period', async () => {
    const actual = await mockCommand.validate({ options: { period: 'abc' } }, commandInfo);
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

  it('get unique device type in teams and export it in a period', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    await mockCommand.action(logger, { options: { period: 'D7' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('produce export using period format and Teams unique device type output in txt', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    await mockCommand.action(logger, { options: { period: 'D7' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('produce export using period format and Teams unique device type output in json', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    await mockCommand.action(logger, { options: { period: 'D7', output: 'json' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('produce export using period format and Teams unique users output in txt', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    await mockCommand.action(logger, { options: { period: 'D7', output: 'text' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('produce export using period format and Teams unique users output in json', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    await mockCommand.action(logger, { options: { period: 'D7' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('produce export using period format and Teams output in json', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period\n2019-08-28,0,0,0,0,0,0,7`);
      }

      return Promise.reject('Invalid request');
    });

    await mockCommand.action(logger, { options: { debug: true, period: 'D7', output: 'json' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(mockCommand.action(logger, { options: { period: 'D7' } } as any), new CommandError('An error has occurred'));
  });
});
