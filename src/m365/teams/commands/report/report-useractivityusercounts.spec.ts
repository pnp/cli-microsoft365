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
import command, { options } from './report-useractivityusercounts.js';

describe(commands.REPORT_USERACTIVITYUSERCOUNTS, () => {
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
    assert.strictEqual(command.name, commands.REPORT_USERACTIVITYUSERCOUNTS);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation on valid \'D7\' period', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D7' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation on valid \'D30\' period', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D30' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation on valid \'D90\' period', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D90' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation on valid \'D180\' period', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D180' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation on invalid period', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if period is missing', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ period: 'D7', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('gets the number of Microsoft Teams users by activity type for the given period', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserCounts(period='D7')`) {
        return `
        Report Refresh Date,Report Date,Team Chat Messages,Private Chat Messages,Calls,Meetings,Other Actions,Report Period
        2019-08-28,2019-08-28,0,0,0,0,0,7
        2019-08-28,2019-08-27,0,0,0,0,0,7
        `;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ period: 'D7' }) });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserCounts(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ period: 'D7' }) }), new CommandError('An error has occurred'));
  });
});
