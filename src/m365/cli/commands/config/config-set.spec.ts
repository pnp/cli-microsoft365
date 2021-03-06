import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import commands from '../../commands';
import configstore from '../../../../configstoreOptions';
const command: Command = require('./config-set');

describe(commands.CONFIG_SET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONFIG_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`sets ${configstore.showHelpOnFailure} property`, (done) => {
    // const options = command.options();

    command.action(logger, {options: {key: configstore.showHelpOnFailure, value: false}}, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it(`sets ${configstore.showHelpOnFailure} property (verbose)`, (done) => {
    // const options = command.options();

    command.action(logger, {options: {key: configstore.showHelpOnFailure, value: false, verbose: true}}, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
        done();
      } catch (e) {
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

  it(`passes validation if service is set to ${configstore.showHelpOnFailure} `, () => {
    const actual = command.validate({ options: { key: configstore.showHelpOnFailure, value: false } });
    assert.strictEqual(actual, true);
  });
});