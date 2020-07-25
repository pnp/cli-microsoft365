import commands from '../../commands';
import Command from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./completion-sh-setup');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import { autocomplete } from '../../../../autocomplete';
import * as chalk from 'chalk';

describe(commands.COMPLETION_SH_SETUP, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let generateShCompletionStub: sinon.SinonStub;
  let setupShCompletionStub: sinon.SinonStub;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    generateShCompletionStub = sinon.stub(autocomplete, 'generateShCompletion').callsFake(() => { });
    setupShCompletionStub = sinon.stub(autocomplete, 'setupShCompletion').callsFake(() => { });
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
    setupShCompletionStub.reset();
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      autocomplete.generateShCompletion,
      autocomplete.setupShCompletion
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_SH_SETUP), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('generates file with commands info', (done) => {
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

  it('sets up command completion in the shell', (done) => {
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(setupShCompletionStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('writes output in verbose mode', (done) => {
    cmdInstance.action({ options: { verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('writes additional info in debug mode', (done) => {
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('Generating command completion...'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});