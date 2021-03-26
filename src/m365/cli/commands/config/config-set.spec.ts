import * as assert from 'assert';
import * as sinon from 'sinon';
import { Cli, Logger } from '../../../../cli';
import Command from '../../../../Command';
import { settingsNames } from '../../../../settingsNames';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./config-set');

describe(commands.CONFIG_SET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.CONFIG_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`sets ${settingsNames.showHelpOnFailure} property`, (done) => {
    const config = Cli.getInstance().config;
    let actualKey: string, actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    command.action(logger, { options: { key: settingsNames.showHelpOnFailure, value: false } }, () => {
      try {
        assert.strictEqual(actualKey, settingsNames.showHelpOnFailure, 'Invalid key');
        assert.strictEqual(actualValue, false, 'Invalid value');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`sets ${settingsNames.output} property to 'text'`, (done) => {
    const output = "text";
    const config = Cli.getInstance().config;
    let actualKey: string, actualValue: any;
    sinon.restore();
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);

    command.action(logger, { options: { key: settingsNames.output, value: output } }, () => {
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

  it(`sets ${settingsNames.output} property to 'json'`, (done) => {
    const output = "json";
    const config = Cli.getInstance().config;
    let actualKey: string, actualValue: any;
    sinon.restore();
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);

    command.action(logger, { options: { key: settingsNames.output, value: output } }, () => {
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

  it('supports specifying key and value', () => {
    const options = command.options();
    let containsOptionKey = false;
    let containsOptionValue = false;
    options.forEach(o => {
      if (o.option.indexOf('--key') > -1) {
        containsOptionKey = true;
      }

      if (o.option.indexOf('--value') > -1) {
        containsOptionValue = true;
      }
    });
    assert(containsOptionKey && containsOptionValue);
  });

  it('fails validation if specified key is invalid ', () => {
    const actual = command.validate({ options: { key: 'invalid', value: false } });
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if service is set to ${settingsNames.showHelpOnFailure} `, () => {
    const actual = command.validate({ options: { key: settingsNames.showHelpOnFailure, value: false } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified output key is invalid ', () => {
    const actual = command.validate({ options: { key: settingsNames.output, value: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });
});