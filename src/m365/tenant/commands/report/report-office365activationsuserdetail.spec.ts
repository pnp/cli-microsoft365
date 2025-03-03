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
import command from './report-office365activationsuserdetail.js';

describe(commands.REPORT_OFFICE365ACTIVATIONSUSERDETAIL, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.REPORT_OFFICE365ACTIVATIONSUSERDETAIL);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets details about users who have activated Microsoft 365', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail`) {
        return `Report Refresh Date,User Principal Name,Display Name,Product Type,Last Activated Date,Windows,Mac,Windows 10 Mobile,iOS,Android,Activated On Shared Computer
        2021-05-25,user1@contoso.onmicrosoft.com,User1,MICROSOFT 365 APPS FOR ENTERPRISE,,0,0,0,0,0,False
        2021-05-25,user1@contoso.onmicrosoft.com,User1,MICROSOFT EXCEL ADVANCED ANALYTICS,,0,0,0,0,0,False`;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('gets details about users who have activated Microsoft 365 (json)', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail`) {
        return `Report Refresh Date,User Principal Name,Display Name,Product Type,Last Activated Date,Windows,Mac,Windows 10 Mobile,iOS,Android,Activated On Shared Computer
        2021-05-25,user1@contoso.onmicrosoft.com,User1,MICROSOFT 365 APPS FOR ENTERPRISE,,0,0,0,0,0,False
        2021-05-25,user1@contoso.onmicrosoft.com,User1,MICROSOFT EXCEL ADVANCED ANALYTICS,,0,0,0,0,0,False`;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { output: 'json' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
    assert(loggerLogSpy.calledWith([{ "Report Refresh Date": "2021-05-25", "User Principal Name": "user1@contoso.onmicrosoft.com", "Display Name": "User1", "Product Type": "MICROSOFT 365 APPS FOR ENTERPRISE", "Last Activated Date": "", "Windows": 0, "Mac": 0, "Windows 10 Mobile": 0, "iOS": 0, "Android": 0, "Activated On Shared Computer": "False" }, { "Report Refresh Date": "2021-05-25", "User Principal Name": "user1@contoso.onmicrosoft.com", "Display Name": "User1", "Product Type": "MICROSOFT EXCEL ADVANCED ANALYTICS", "Last Activated Date": "", "Windows": 0, "Mac": 0, "Windows 10 Mobile": 0, "iOS": 0, "Android": 0, "Activated On Shared Computer": "False" }]));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
