import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./completion-sh-update');

describe(commands.COMPLETION_SH_UPDATE, () => {
  let log: string[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;
  let generateShCompletionStub: sinon.SinonStub;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    generateShCompletionStub = sinon.stub(autocomplete, 'generateShCompletion').callsFake(() => { });
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    generateShCompletionStub.reset();
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      autocomplete.generateShCompletion
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_SH_UPDATE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('builds command completion', (done) => {
    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(generateShCompletionStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('build command completion (debug)', (done) => {
    command.action(logger, { options: { debug: true } }, () => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});