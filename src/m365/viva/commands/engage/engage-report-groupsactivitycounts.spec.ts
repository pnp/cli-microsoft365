import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './engage-report-groupsactivitycounts.js';
import yammerCommands from './yammerCommands.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.ENGAGE_REPORT_GROUPSACTIVITYCOUNTS, () => {
  let log: string[];
  let logger: Logger;

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
    assert.strictEqual(command.name, commands.ENGAGE_REPORT_GROUPSACTIVITYCOUNTS);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [yammerCommands.REPORT_GROUPSACTIVITYCOUNTS]);
  });

  it('correctly logs deprecation warning for yammer command', async () => {
    const chalk = (await import('chalk')).default;
    const loggerErrSpy = sinon.spy(logger, 'logToStderr');
    const commandNameStub = sinon.stub(cli, 'currentCommandName').value(yammerCommands.REPORT_GROUPSACTIVITYCOUNTS);
    sinon.stub(request, 'get').resolves('Report Refresh Date,User Principal Name,Display Name,User State,State Change Date,Last Activity Date,Used Web,Used Windows Phone,Used Android Phone,Used iPhone,Used iPad,Used Others,Report Period');

    await command.action(logger, { options: { period: 'D7' } });
    assert.deepStrictEqual(loggerErrSpy.firstCall.firstArg, chalk.yellow(`Command '${yammerCommands.REPORT_GROUPSACTIVITYCOUNTS}' is deprecated. Please use '${commands.ENGAGE_REPORT_GROUPSACTIVITYCOUNTS}' instead.`));

    sinonUtil.restore([loggerErrSpy, commandNameStub]);
  });

  it('gets the report for the last week', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getYammerGroupsActivityCounts(period='D7')`) {
        return `Report Refresh Date,Liked,Posted,Read,Report Date,Report Period`;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { period: 'D7' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getYammerGroupsActivityCounts(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });
});
