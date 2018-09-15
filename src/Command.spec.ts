import * as sinon from 'sinon';
import * as assert from 'assert';
import Command, {
  CommandValidate,
  CommandCancel,
  CommandOption,
  CommandTypes,
  CommandError
} from './Command';
import Utils from './Utils';
import appInsights from './appInsights';
const vorpal: Vorpal = require('./vorpal-init');

class MockCommand1 extends Command {
  public get name(): string {
    return 'Mock command';
  }

  public get description(): string {
    return 'Mock command description';
  }

  public alias(): string[] | undefined {
    return ['mc1'];
  }

  public autocomplete(): string[] | undefined {
    const autocomplete = ['param1', 'param2'];

    const parentAutocomplete: string[] | undefined = super.autocomplete();
    if (parentAutocomplete) {
      return autocomplete.concat(parentAutocomplete);
    }
    else {
      return autocomplete;
    }
  }

  public allowUnknownOptions(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    this.showDeprecationWarning(cmd, 'mc1', this.name);

    appInsights.trackEvent({
      name: this.getUsedCommandName(cmd)
    });

    cb();
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }

  public validate(): CommandValidate | undefined {
    return () => {
      return true;
    };
  }

  public cancel(): CommandCancel | undefined {
    return () => { };
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['']
    };
  }

  public options(): CommandOption[] {
    return [{
      option: '--debug',
      description: 'Runs command with debug logging'
    }];
  }
}

class MockCommand2 extends Command {
  public get name(): string {
    return 'Mock command 2 [opt]';
  }

  public get description(): string {
    return 'Mock command 2 description';
  }

  public commandAction(): void {
  }

  public commandHelp(args: any, log: (message: string) => void): void {
    log('MockCommand2 help');
  }

  public handlePromiseError(response: any, cmd: CommandInstance, callback: (err?: any) => void): void {
    this.handleRejectedODataJsonPromise(response, cmd, callback);
  }
}

describe('Command', () => {
  const vcmd = {
    action: () => vcmd,
    alias: () => vcmd,
    option: () => vcmd,
    validate: () => vcmd,
    cancel: () => vcmd,
    help: () => vcmd,
    types: () => vcmd,
    allowUnknownOptions: () => vcmd
  };
  let actionSpy: sinon.SinonSpy;
  let aliasSpy: sinon.SinonSpy;
  let optionSpy: sinon.SinonSpy;
  let validateSpy: sinon.SinonSpy;
  let cancelSpy: sinon.SinonSpy;
  let helpSpy: sinon.SinonSpy;
  let typesSpy: sinon.SinonSpy;
  let telemetry: any;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    actionSpy = sinon.spy(vcmd, 'action');
    aliasSpy = sinon.spy(vcmd, 'alias');
    optionSpy = sinon.spy(vcmd, 'option');
    validateSpy = sinon.spy(vcmd, 'validate');
    cancelSpy = sinon.spy(vcmd, 'cancel');
    helpSpy = sinon.spy(vcmd, 'help');
    typesSpy = sinon.spy(vcmd, 'types');
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vcmd.action,
      vcmd.alias,
      vcmd.option,
      vcmd.validate,
      vcmd.cancel,
      vcmd.help,
      vcmd.types,
      vcmd.allowUnknownOptions
    ]);
  });

  after(() => {
    Utils.restore(appInsights.trackEvent);
  })

  it('initiates command 1 with vorpal', () => {
    const cmd = new MockCommand1();
    const vorpalCommandStub = sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(vorpalCommandStub.calledOnce);
  });

  it('initiates command 2 with vorpal', () => {
    const cmd = new MockCommand2();
    const vorpalCommandStub = sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(vorpalCommandStub.calledOnce);
  });

  it('initiates command with command name', () => {
    const cmd = new MockCommand1();
    let name;
    sinon.stub(vorpal, 'command').callsFake((_name) => {
      name = _name;
      return vcmd;
    });
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert.equal(name, cmd.name);
  });

  it('initiates command with command description', () => {
    const cmd = new MockCommand1();
    let description;
    sinon.stub(vorpal, 'command').callsFake((_name, _description) => {
      description = _description;
      return vcmd;
    });
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert.equal(description, cmd.description);
  });

  it('initiates command with command autocomplete', () => {
    const cmd = new MockCommand1();
    let autocomplete;
    sinon.stub(vorpal, 'command').callsFake((_name, _description, _autocomplete) => {
      autocomplete = _autocomplete;
      return vcmd;
    });
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert.deepEqual(autocomplete, cmd.autocomplete());
  });

  it('configures command action', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(actionSpy.calledOnce);
  });

  it('configures options when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(optionSpy.calledOnce);
  });

  it('configures alias when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(aliasSpy.calledOnce);
  });

  it('configures validation when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(validateSpy.calledOnce);
  });

  it('doesn\'t configure validation when unavailable', () => {
    const cmd = new MockCommand2();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(validateSpy.notCalled);
  });

  it('configures cancellation when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(cancelSpy.calledOnce);
  });

  it('doesn\'t configure cancellation when unavailable', () => {
    const cmd = new MockCommand2();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(cancelSpy.notCalled);
  });

  it('configures help when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(helpSpy.calledOnce);
  });

  it('configures types when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(typesSpy.calledOnce);
  });

  it('doesn\'t configure type when unavailable', () => {
    const cmd = new MockCommand2();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(typesSpy.notCalled);
  });

  it('returns command name without arguments', () => {
    const cmd = new MockCommand2();
    assert.equal(cmd.getCommandName(), 'Mock command 2');
  });

  it('prints help using the log argument when called from the help command', () => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(vorpal, '_command').value({
      command: 'help mock2'
    });
    const log = (msg?: string) => { };
    const logSpy = sinon.spy(log);
    const mock = new MockCommand2();
    const cmd = {
      help: mock.help()
    };
    cmd.help({}, logSpy);
    sandbox.restore();
    assert(logSpy.called);
  });

  it('displays error message when it\'s serialized in the error property', () => {
    const cmd = {
      commandWrapper: {
        command: 'command'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand2();
    mock.handlePromiseError({
      error: JSON.stringify({
        error: {
          message: 'An error has occurred'
        }
      })
    }, cmd, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
    });
  });

  it('displays the raw error message when the serialized value from the error property is not an error object', () => {
    const cmd = {
      commandWrapper: {
        command: 'command'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand2();
    mock.handlePromiseError({
      error: JSON.stringify({
        error: {
          id: '123'
        }
      })
    }, cmd, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(JSON.stringify({
        error: {
          id: '123'
        }
      }))));
    });
  });

  it('displays the raw error message when the serialized value from the error property is not a JSON object', () => {
    const cmd = {
      commandWrapper: {
        command: 'command'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand2();
    mock.handlePromiseError({
      error: 'abc'
    }, cmd, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
    });
  });

  it('displays error message coming from ADALJS', () => {
    const cmd = {
      commandWrapper: {
        command: 'command'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand2();
    mock.handlePromiseError({
      error: { error_description: 'abc' }
    }, cmd, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
    });
  });

  it('shows deprecation warning when command executed using the deprecated name', () => {
    const cmd = {
      commandWrapper: {
        command: 'mc1'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    const cmdLogSpy: sinon.SinonSpy = sinon.spy(cmd, 'log');
    const mock = new MockCommand1();
    mock.commandAction(cmd, {}, (err?: any): void => {
      assert(cmdLogSpy.calledWith(vorpal.chalk.yellow(`Command 'mc1' is deprecated. Please use 'Mock command' instead`)))
    });
  });

  it('logs command name in the telemetry when command name used', (done) => {
    const cmd = {
      commandWrapper: {
        command: 'Mock command'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand1();
    mock.commandAction(cmd, {}, (err?: any): void => {
      try {
        assert.equal(telemetry.name, 'Mock command');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs command alias in the telemetry when command alias used', (done) => {
    const cmd = {
      commandWrapper: {
        command: 'mc1'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand1();
    mock.commandAction(cmd, {}, (err?: any): void => {
      try {
        assert.equal(telemetry.name, 'mc1');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs empty command name in telemetry when command called using something else than name or alias', (done) => {
    const cmd = {
      commandWrapper: {
        command: 'foo'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand1();
    mock.commandAction(cmd, {}, (err?: any): void => {
      try {
        assert.equal(telemetry.name, '');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});