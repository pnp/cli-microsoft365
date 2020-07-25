import commands from '../../commands';
import Command from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./completion-sh-update');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import { autocomplete } from '../../../../autocomplete';
import * as chalk from 'chalk';

describe(commands.COMPLETION_SH_UPDATE, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let generateShCompletionStub: sinon.SinonStub;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    generateShCompletionStub = sinon.stub(autocomplete, 'generateShCompletion').callsFake(() => { });
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
    cmdInstance.action({ options: { debug: false } }, () => {
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
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});