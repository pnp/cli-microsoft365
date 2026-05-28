import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
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
import command, { options } from './report-usageaccountdetail.js';

describe(commands.REPORT_USAGEACCOUNTDETAIL, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodType;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    assert.strictEqual(command.name, commands.REPORT_USAGEACCOUNTDETAIL);
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

  it('gets the report for the last week', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D7')`) {
        return `Report Refresh Date,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Storage Used (Byte),Storage Allocated (Byte),Owner Principal Name,Report Period`;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: options.parse({ period: 'D7' }) });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('gets the report for the given date', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(date=2019-07-13)`) {
        return `Report Refresh Date,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Storage Used (Byte),Storage Allocated (Byte),Owner Principal Name,Report Period`;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: options.parse({ date: '2019-07-13' }) });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(date=2019-07-13)");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: options.parse({ period: 'D7' }) }), new CommandError('An error has occurred'));
  });
});
