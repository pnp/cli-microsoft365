import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./report-office365activationsuserdetail');

describe(commands.REPORT_OFFICE365ACTIVATIONSUSERDETAIL, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent,
      pid.getProcessName,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.REPORT_OFFICE365ACTIVATIONSUSERDETAIL), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets details about users who have activated Microsoft 365', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail`) {
        return Promise.resolve(`Report Refresh Date,User Principal Name,Display Name,Product Type,Last Activated Date,Windows,Mac,Windows 10 Mobile,iOS,Android,Activated On Shared Computer
        2021-05-25,user1@contoso.onmicrosoft.com,User1,MICROSOFT 365 APPS FOR ENTERPRISE,,0,0,0,0,0,False
        2021-05-25,user1@contoso.onmicrosoft.com,User1,MICROSOFT EXCEL ADVANCED ANALYTICS,,0,0,0,0,0,False`);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: {} });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('gets details about users who have activated Microsoft 365 (json)', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail`) {
        return Promise.resolve(`Report Refresh Date,User Principal Name,Display Name,Product Type,Last Activated Date,Windows,Mac,Windows 10 Mobile,iOS,Android,Activated On Shared Computer
        2021-05-25,user1@contoso.onmicrosoft.com,User1,MICROSOFT 365 APPS FOR ENTERPRISE,,0,0,0,0,0,False
        2021-05-25,user1@contoso.onmicrosoft.com,User1,MICROSOFT EXCEL ADVANCED ANALYTICS,,0,0,0,0,0,False`);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { output: 'json' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
    assert(loggerLogSpy.calledWith([{ "Report Refresh Date": "2021-05-25", "User Principal Name": "user1@contoso.onmicrosoft.com", "Display Name": "User1", "Product Type": "MICROSOFT 365 APPS FOR ENTERPRISE", "Last Activated Date": "", "Windows": 0, "Mac": 0, "Windows 10 Mobile": 0, "iOS": 0, "Android": 0, "Activated On Shared Computer": "False" }, { "Report Refresh Date": "2021-05-25", "User Principal Name": "user1@contoso.onmicrosoft.com", "Display Name": "User1", "Product Type": "MICROSOFT EXCEL ADVANCED ANALYTICS", "Last Activated Date": "", "Windows": 0, "Mac": 0, "Windows 10 Mobile": 0, "iOS": 0, "Android": 0, "Activated On Shared Computer": "False" }]));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});