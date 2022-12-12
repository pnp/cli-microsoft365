import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./report-office365activationcounts');

describe(commands.REPORT_OFFICE365ACTIVATIONCOUNTS, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.REPORT_OFFICE365ACTIVATIONCOUNTS), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets the count of Microsoft 365 activations', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365ActivationCounts`) {
        return Promise.resolve(`Report Refresh Date,Product Type,Windows,Mac,Android,iOS,Windows 10 Mobile
        2021-05-24,MICROSOFT 365 APPS FOR ENTERPRISE,2,0,0,0,0
        2021-05-24,MICROSOFT EXCEL ADVANCED ANALYTICS,0,0,0,0,0`);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: {} });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationCounts");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('gets the count of Microsoft 365 activations (json)', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365ActivationCounts`) {
        return Promise.resolve(`Report Refresh Date,Product Type,Windows,Mac,Android,iOS,Windows 10 Mobile
        2021-05-24,MICROSOFT 365 APPS FOR ENTERPRISE,2,0,0,0,0
        2021-05-24,MICROSOFT EXCEL ADVANCED ANALYTICS,0,0,0,0,0`);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { output: 'json' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationCounts");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
    assert(loggerLogSpy.calledWith([{ "Report Refresh Date": "2021-05-24", "Product Type": "MICROSOFT 365 APPS FOR ENTERPRISE", "Windows": 2, "Mac": 0, "Android": 0, "iOS": 0, "Windows 10 Mobile": 0 }, { "Report Refresh Date": "2021-05-24", "Product Type": "MICROSOFT EXCEL ADVANCED ANALYTICS", "Windows": 0, "Mac": 0, "Android": 0, "iOS": 0, "Windows 10 Mobile": 0 }]));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
