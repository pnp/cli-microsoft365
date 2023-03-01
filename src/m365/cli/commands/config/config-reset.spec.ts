import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import { settingsNames } from '../../../../settingsNames';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./config-reset');

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
  });

  after(() => {
    sinonUtil.restore([
      Cli.getInstance().config.set,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
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
