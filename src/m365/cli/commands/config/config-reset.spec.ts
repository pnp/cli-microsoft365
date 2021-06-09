import * as assert from 'assert';
import * as sinon from 'sinon';
import { Cli, Logger } from '../../../../cli';
import Command from '../../../../Command';
import { settingsNames } from '../../../../settingsNames';
import Utils from '../../../../Utils';
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
    Utils.restore(Cli.getInstance().config.set);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONFIG_RESET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`Resets ${settingsNames.showHelpOnFailure} property`, (done) => {
    const config = Cli.getInstance().config;
    let actualKey: string, actualValue: string;

    sinon.restore();
    sinon.stub(config, 'set').callsFake(((key: string) => {
      actualKey = key;
      actualValue = 'true';
    }) as any);

    command.action(logger, { options: { key: settingsNames.showHelpOnFailure } }, () => {
      try {
        assert.strictEqual(actualKey, settingsNames.showHelpOnFailure, 'Invalid key');
        assert.strictEqual(actualValue, 'true', 'Invalid value');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`Resets ${settingsNames.printErrorsAsPlainText} property`, (done) => {
    const config = Cli.getInstance().config;
    let actualKey: string, actualValue: string;

    sinon.restore();
    sinon.stub(config, 'set').callsFake(((key: string) => {
      actualKey = key;
      actualValue = 'true';
    }) as any);

    command.action(logger, { options: { key: settingsNames.printErrorsAsPlainText } }, () => {
      try {
        assert.strictEqual(actualKey, settingsNames.printErrorsAsPlainText, 'Invalid key');
        assert.strictEqual(actualValue, 'true', 'Invalid value');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`Resets ${settingsNames.output} property`, (done) => {
    const output = "text";
    const config = Cli.getInstance().config;
    let actualKey: string, actualValue: string;

    sinon.restore();
    sinon.stub(config, 'set').callsFake(((key: string) => {
      actualKey = key;
      actualValue = 'text';
    }) as any);

    command.action(logger, { options: { key: settingsNames.output } }, () => {
      try {
        assert.strictEqual(actualKey, settingsNames.output, 'Invalid key');
        assert.strictEqual(actualValue, output, 'Invalid value');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`Resets ${settingsNames.errorOutput} property`, (done) => {
    const output = "stderr";
    const config = Cli.getInstance().config;
    let actualKey: string, actualValue: string;

    sinon.restore();
    sinon.stub(config, 'set').callsFake(((key: string) => {
      actualKey = key;
      actualValue = 'stderr';
    }) as any);

    command.action(logger, { options: { key: settingsNames.errorOutput } }, () => {
      try {
        assert.strictEqual(actualKey, settingsNames.errorOutput, 'Invalid key');
        assert.strictEqual(actualValue, output, 'Invalid value');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`Resets all configuration settings to default`, (done) => {
    const config = Cli.getInstance().config;
    let errorOutputKey: string, errorOutputValue: string
      , outputKey: string, outputValue: string
      , printErrorsAsPlainTextKey: string, printErrorsAsPlainTextValue: string
      , showHelpOnFailureKey: string, showHelpOnFailureValue: string;

    sinon.restore();

    sinon.stub(config, 'set').callsFake((() => {
      errorOutputKey = settingsNames.errorOutput;
      errorOutputValue = 'stderr';
      outputKey = settingsNames.output;
      outputValue = 'text';
      printErrorsAsPlainTextKey = settingsNames.printErrorsAsPlainText;
      printErrorsAsPlainTextValue = 'true';
      showHelpOnFailureKey = settingsNames.showHelpOnFailure;
      showHelpOnFailureValue = 'true';
    }) as any);

    command.action(logger, { options: {} }, () => {
      try {
        assert.strictEqual(errorOutputKey, settingsNames.errorOutput, 'Invalid key');
        assert.strictEqual(errorOutputValue, 'stderr', 'Invalid value');

        assert.strictEqual(outputKey, settingsNames.output, 'Invalid key');
        assert.strictEqual(outputValue, 'text', 'Invalid value');

        assert.strictEqual(printErrorsAsPlainTextKey, settingsNames.printErrorsAsPlainText, 'Invalid key');
        assert.strictEqual(printErrorsAsPlainTextValue, 'true', 'Invalid value');

        assert.strictEqual(showHelpOnFailureKey, settingsNames.showHelpOnFailure, 'Invalid key');
        assert.strictEqual(showHelpOnFailureValue, 'true', 'Invalid value');
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