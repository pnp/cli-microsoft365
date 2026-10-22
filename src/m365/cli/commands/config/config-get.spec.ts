import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { settingsNames } from '../../../../settingsNames.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './config-get.js';

describe(commands.CONFIG_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    sinonUtil.restore(cli.getConfig().get);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONFIG_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`gets value of the specified property`, async () => {
    const config = cli.getConfig();
    sinon.stub(config, 'get').callsFake(_ => 'json');
    await command.action(logger, { options: { key: settingsNames.output } });
    assert(loggerSpy.calledWith('json'));
  });

  it(`returns undefined if the specified setting is not configured`, async () => {
    const config = cli.getConfig();
    sinon.stub(config, 'get').callsFake(_ => undefined);
    await command.action(logger, { options: { key: settingsNames.output } });
    assert(loggerSpy.calledWith(undefined));
  });

  it('fails validation if specified key is invalid ', () => {
    const actual = commandOptionsSchema.safeParse({ key: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it(`passes validation if setting is set to ${settingsNames.showHelpOnFailure}`, () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.showHelpOnFailure });
    assert.strictEqual(actual.success, true);
  });
});
