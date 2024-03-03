import assert from 'assert';
import sinon from 'sinon';
import { telemetry } from '../../../../telemetry.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './config-list.js';

describe(commands.CONFIG_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    loggerSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore(cli.getConfig().all);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONFIG_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('returns a list of all the self set properties', async () => {
    const config = cli.getConfig();
    sinon.stub(config, 'all').value({ 'errorOutput': 'stdout' });

    await command.action(logger, { options: {} });
    assert(loggerSpy.calledWith({ 'errorOutput': 'stdout' }));
  });
});
