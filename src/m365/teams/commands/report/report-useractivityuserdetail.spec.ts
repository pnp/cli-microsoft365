import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './report-useractivityuserdetail.js';

describe(commands.REPORT_USERACTIVITYUSERDETAIL, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    assert.strictEqual(command.name, commands.REPORT_USERACTIVITYUSERDETAIL);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation on valid \'D7\' period', async () => {
    const actual = await commandOptionsSchema.safeParseAsync({ period: 'D7' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation on valid date', async () => {
    const actual = await commandOptionsSchema.safeParseAsync({ date: '2019-07-13' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation on invalid period', async () => {
    const actual = await commandOptionsSchema.safeParseAsync({ period: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither period nor date is specified', async () => {
    const actual = await commandOptionsSchema.safeParseAsync({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both period and date options set', async () => {
    const actual = await commandOptionsSchema.safeParseAsync({ period: 'D7', date: '2019-07-13' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation on invalid date format', async () => {
    const actual = await commandOptionsSchema.safeParseAsync({ date: '10.10.2019' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', async () => {
    const actual = await commandOptionsSchema.safeParseAsync({ period: 'D7', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('gets details about Microsoft Teams user activity by user for the given date', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(date=2019-07-13)`) {
        return `
        Report Refresh Date,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Assigned Products,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Has Other Action,Report Period
        2019-08-14,abisha@contoso.onmicrosoft.com,,False,,,0,0,0,0,No,7
        2019-08-14,same@contoso.onmicrosoft.com,2019-05-22,False,,OFFICE 365 E3 DEVELOPER+MICROSOFT FLOW FREE,0,0,0,0,No,7
        `;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ date: '2019-07-13' }) });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(date=2019-07-13)");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('gets details about Microsoft Teams user activity by user for the given period', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D7')`) {
        return `
        Report Refresh Date,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Assigned Products,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Has Other Action,Report Period
        2019-08-14,abisha@contoso.onmicrosoft.com,,False,,,0,0,0,0,No,7
        `;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ period: 'D7' }) });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ period: 'D7' }) }), new CommandError('An error has occurred'));
  });
});
