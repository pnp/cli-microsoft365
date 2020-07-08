import commands from '../../commands';
import Command from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./completion-sh-setup');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import { autocomplete } from '../../../../autocomplete';

describe(commands.COMPLETION_SH_SETUP, () => {
  let vorpal: Vorpal;
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
    vorpal = require('../../../../vorpal-init');
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
    Utils.restore([
      vorpal.find
    ]);
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
    assert.equal(command.name.startsWith(commands.COMPLETION_SH_SETUP), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.COMPLETION_SH_SETUP));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});