import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './report-office365activationsusercounts.js';

describe(commands.REPORT_OFFICE365ACTIVATIONSUSERCOUNTS, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.REPORT_OFFICE365ACTIVATIONSUSERCOUNTS);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets details of office 365 subscription user counts', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserCounts`) {
        return `Report Refresh Date,Product Type,Assigned,Activated,Shared Computer Activation
        2021-05-24,MICROSOFT 365 APPS FOR ENTERPRISE,3,2,0
        2021-05-24,MICROSOFT EXCEL ADVANCED ANALYTICS,3,0,0`;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserCounts");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('gets details of office 365 subscription user counts (json)', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserCounts`) {
        return `Report Refresh Date,Product Type,Assigned,Activated,Shared Computer Activation
        2021-05-24,MICROSOFT 365 APPS FOR ENTERPRISE,3,2,0
        2021-05-24,MICROSOFT EXCEL ADVANCED ANALYTICS,3,0,0`;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { output: 'json' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserCounts");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
    assert(loggerLogSpy.calledWith([{ "Report Refresh Date": "2021-05-24", "Product Type": "MICROSOFT 365 APPS FOR ENTERPRISE", "Assigned": 3, "Activated": 2, "Shared Computer Activation": 0 }, { "Report Refresh Date": "2021-05-24", "Product Type": "MICROSOFT EXCEL ADVANCED ANALYTICS", "Assigned": 3, "Activated": 0, "Shared Computer Activation": 0 }]));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
