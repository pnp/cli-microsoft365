import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./report-mailboxusagedetail');

describe(commands.REPORT_MAILBOXUSAGEDETAIL, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.REPORT_MAILBOXUSAGEDETAIL), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets the report for the last week', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D7')`) {
        return Promise.resolve(`
          Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Created Date,Last Activity Date,Item Count,Storage Used (Byte),Issue Warning Quota (Byte),Prohibit Send Quota (Byte),Prohibit Send/Receive Quota (Byte),Deleted Item Count,Deleted Item Size (Byte),Report Period`
        );
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, period: 'D7' } }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D7')");
        assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});