import * as assert from 'assert';
import * as sinon from 'sinon';
import { Cli, Logger } from '../../../../cli';
import Command from '../../../../Command';
import { settingsNames } from '../../../../settingsNames';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./config-reset');

describe(commands.CONFIG_RESET, () => {
  let log: any[];
  let logger: Logger;

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
    sinonUtil.restore(Cli.getInstance().config.set);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONFIG_RESET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`Resets a specific configuration option to its default value`, (done) => {
    const output = undefined;
    const config = Cli.getInstance().config;

    let actualKey: string, actualValue: any;

    sinon.restore();
    sinon.stub(config, 'delete').callsFake(((key: string) => {
      actualKey = key;
      actualValue = undefined;
    }) as any);

    command.action(logger, { options: { key: settingsNames.output, value: output } }, () => {
      try {
        assert.strictEqual(actualKey, settingsNames.output, 'Invalid key');
        assert.strictEqual(actualValue, undefined, 'Invalid value');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`Resets all configuration settings to default`, (done) => {
    const config = Cli.getInstance().config;
    let errorOutputKey: string, errorOutputValue: any
      , outputKey: string, outputValue: any
      , printErrorsAsPlainTextKey: string, printErrorsAsPlainTextValue: any
      , showHelpOnFailureKey: string, showHelpOnFailureValue: any;

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

    command.action(logger, { options: {} }, () => {
      try {
        assert.strictEqual(errorOutputKey, settingsNames.errorOutput, 'Invalid key');
        assert.strictEqual(errorOutputValue, undefined, 'Invalid value');

        assert.strictEqual(outputKey, settingsNames.output, 'Invalid key');
        assert.strictEqual(outputValue, undefined, 'Invalid value');

        assert.strictEqual(printErrorsAsPlainTextKey, settingsNames.printErrorsAsPlainText, 'Invalid key');
        assert.strictEqual(printErrorsAsPlainTextValue, undefined, 'Invalid value');

        assert.strictEqual(showHelpOnFailureKey, settingsNames.showHelpOnFailure, 'Invalid key');
        assert.strictEqual(showHelpOnFailureValue, undefined, 'Invalid value');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if specified key is invalid', () => {
    const actual = command.validate({ options: { key: 'invalid', value: false } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if key is not specified', () => {
    const actual = command.validate({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('supports specifying key', () => {
    const options = command.options();
    let containsOptionKey = false;
    options.forEach(o => {
      if (o.option.indexOf('--key') > -1) {
        containsOptionKey = true;
      }
    });

    assert(containsOptionKey);
  });
});