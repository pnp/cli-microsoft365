import assert from 'assert';
import sinon from 'sinon';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { settingsNames } from '../../../../settingsNames.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import commands from '../../commands.js';
import command from './config-reset.js';

describe(commands.CONFIG_RESET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    commandInfo = Cli.getCommandInfo(command);
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONFIG_RESET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`Resets a specific configuration option to its default value`, async () => {
    const output = undefined;
    const config = Cli.getInstance().config;

    let actualKey: string = '', actualValue: any;

    sinon.restore();
    sinon.stub(config, 'delete').callsFake(((key: string) => {
      actualKey = key;
      actualValue = undefined;
    }) as any);

    await command.action(logger, { options: { key: settingsNames.output, value: output } });
    assert.strictEqual(actualKey, settingsNames.output, 'Invalid key');
    assert.strictEqual(actualValue, undefined, 'Invalid value');
  });

  it(`Resets all configuration settings to default`, async () => {
    const config = Cli.getInstance().config;
    let errorOutputKey: string = '', errorOutputValue: any
      , outputKey: string = '', outputValue: any
      , printErrorsAsPlainTextKey: string = '', printErrorsAsPlainTextValue: any
      , showHelpOnFailureKey: string = '', showHelpOnFailureValue: any;

    sinon.restore();

    sinon.stub(config, 'clear').callsFake((() => {
      errorOutputKey = settingsNames.errorOutput;
      errorOutputValue = undefined;

      outputKey = settingsNames.output;
      outputValue = undefined;

      printErrorsAsPlainTextKey = settingsNames.printErrorsAsPlainText;
      printErrorsAsPlainTextValue = undefined;

      showHelpOnFailureKey = settingsNames.showHelpOnFailure;
      showHelpOnFailureValue = undefined;
    }) as any);

    await command.action(logger, { options: {} });
    assert.strictEqual(errorOutputKey, settingsNames.errorOutput, 'Invalid key');
    assert.strictEqual(errorOutputValue, undefined, 'Invalid value');

    assert.strictEqual(outputKey, settingsNames.output, 'Invalid key');
    assert.strictEqual(outputValue, undefined, 'Invalid value');

    assert.strictEqual(printErrorsAsPlainTextKey, settingsNames.printErrorsAsPlainText, 'Invalid key');
    assert.strictEqual(printErrorsAsPlainTextValue, undefined, 'Invalid value');

    assert.strictEqual(showHelpOnFailureKey, settingsNames.showHelpOnFailure, 'Invalid key');
    assert.strictEqual(showHelpOnFailureValue, undefined, 'Invalid value');
  });

  it('fails validation if specified key is invalid', async () => {
    const actual = await command.validate({ options: { key: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if key is not specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying key', () => {
    const options = command.options;
    let containsOptionKey = false;
    options.forEach(o => {
      if (o.option.indexOf('--key') > -1) {
        containsOptionKey = true;
      }
    });

    assert(containsOptionKey);
  });
});
