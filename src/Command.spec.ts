import * as sinon from 'sinon';
import * as assert from 'assert';
import auth from './Auth';
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
    return 'mock-command';
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
      string: ['option2']
    };
  }

  public options(): CommandOption[] {
    return [
      {
        option: '--debug',
        description: 'Runs command with debug logging'
      },
      {
        option: '--option1 [option1]',
        description: 'Some option'
      },
      {
        option: '--option2 [option2]',
        description: 'Some other option'
      }
    ];
  }

  public trackUnknownOptionsPublic(telemetryProps: any, options: any) {
    return this.trackUnknownOptions(telemetryProps, options);
  }

  public addUnknownOptionsToPayloadPublic(payload: any, options: any) {
    return this.addUnknownOptionsToPayload(payload, options);
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

class MockCommand3 extends Command {
  public get name(): string {
    return 'mock-command';
  }

  public get description(): string {
    return 'Mock command description';
  }

  public commandAction(): void {
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }

  public options(): CommandOption[] {
    return [
      {
        option: '--debug',
        description: 'Runs command with debug logging'
      },
      {
        option: '--option1 [option1]',
        description: 'Some option'
      }
    ];
  }
}

class MockCommand4 extends Command {
  public get name(): string {
    return 'mock-command';
  }

  public get description(): string {
    return 'Mock command description';
  }

  public allowUnknownOptions(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    cb();
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }

  public options(): CommandOption[] {
    return [
      {
        option: '--debug',
        description: 'Runs command with debug logging'
      }
    ];
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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
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
    auth.service.connected = true;
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
      vcmd.allowUnknownOptions,
      vorpal.command,
      process.exit,
      vorpal.util.parseCommand
    ]);
    vorpal.commands = [];
    (vorpal as any)._command = undefined;
    auth.service.connected = false;
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
  })

  it('initiates command 1 with vorpal', () => {
    const cmd = new MockCommand1();
    const vorpalCommandStub = sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(vorpalCommandStub.calledOnce);
  });

  it('initiates command 2 with vorpal', () => {
    const cmd = new MockCommand2();
    const vorpalCommandStub = sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(vorpalCommandStub.calledOnce);
  });

  it('initiates command with command name', () => {
    const cmd = new MockCommand1();
    let name;
    sinon.stub(vorpal, 'command').callsFake((_name) => {
      name = _name;
      return vcmd as any;
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
      return vcmd as any;
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
      return vcmd as any;
    });
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert.deepEqual(autocomplete, cmd.autocomplete());
  });

  it('configures command action', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(actionSpy.calledOnce);
  });

  it('configures options when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(optionSpy.calledThrice); // there are three options
  });

  it('configures alias when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(aliasSpy.calledOnce);
  });

  it('configures validation when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(validateSpy.calledOnce);
  });

  it('doesn\'t configure validation when unavailable', () => {
    const cmd = new MockCommand2();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(validateSpy.notCalled);
  });

  it('configures cancellation when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(cancelSpy.calledOnce);
  });

  it('doesn\'t configure cancellation when unavailable', () => {
    const cmd = new MockCommand2();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(cancelSpy.notCalled);
  });

  it('configures help when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(helpSpy.calledOnce);
  });

  it('configures types when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(typesSpy.calledOnce);
  });

  it('doesn\'t configure type when unavailable', () => {
    const cmd = new MockCommand2();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd as any);
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
      assert(cmdLogSpy.calledWith(vorpal.chalk.yellow(`Command 'mc1' is deprecated. Please use 'mock-command' instead`)))
    });
  });

  it('logs command name in the telemetry when command name used', (done) => {
    const mock = new MockCommand1();
    const cmd = {
      action: mock.action(),
      commandWrapper: {
        command: 'mock-command'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    cmd.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, 'mock-command');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs command alias in the telemetry when command alias used', (done) => {
    const mock = new MockCommand1();
    const cmd = {
      action: mock.action(),
      commandWrapper: {
        command: 'mc1'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    cmd.action({ options: {} }, () => {
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
    const mock = new MockCommand1();
    const cmd = {
      action: mock.action(),
      commandWrapper: {
        command: 'foo'
      },
      log: (msg?: string) => { },
      prompt: () => { }
    };
    cmd.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, '');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t remove leading zeroes from unknown options', (done) => {
    const cmd = new MockCommand1();
    const delimiter = (vorpal as any)._delimiter;
    const argv = process.argv;
    vorpal.delimiter('');
    sinon.stub(cmd as any, 'initAction').callsFake((args) => {
      try {
        assert.strictEqual(args.options.option3, '00123');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        vorpal.delimiter(delimiter);
        process.argv = argv;
      }
    });
    sinon.stub(process, 'exit');
    cmd.init(vorpal);
    process.argv = ['node', 'm365', 'mock-command', '--option3', '00123'];
    vorpal.parse(['node', 'm365', 'mock-command', '--option3', '00123']);
  });

  it('removes leading zeroes from known options that aren\'t a string', (done) => {
    const cmd = new MockCommand1();
    const delimiter = (vorpal as any)._delimiter;
    const argv = process.argv;
    vorpal.delimiter('');
    sinon.stub(cmd as any, 'initAction').callsFake((args) => {
      try {
        assert.strictEqual(args.options.option1, 123);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        vorpal.delimiter(delimiter);
        process.argv = argv;
      }
    });
    sinon.stub(process, 'exit');
    cmd.init(vorpal);
    process.argv = ['node', 'm365', 'mock-command', '--option1', '00123'];
    vorpal.parse(['node', 'm365', 'mock-command', '--option1', '00123']);
  });

  it('doesn\'t remove leading zeroes from known options that are a string', (done) => {
    const cmd = new MockCommand1();
    const delimiter = (vorpal as any)._delimiter;
    const argv = process.argv;
    vorpal.delimiter('');
    sinon.stub(cmd as any, 'initAction').callsFake((args) => {
      try {
        assert.strictEqual(args.options.option2, '00123');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        vorpal.delimiter(delimiter);
        process.argv = argv;
      }
    });
    sinon.stub(process, 'exit');
    cmd.init(vorpal);
    process.argv = ['node', 'm365', 'mock-command', '--option2', '00123'];
    vorpal.parse(['node', 'm365', 'mock-command', '--option2', '00123']);
  });

  it('doesn\'t remove leading zeroes from unknown options where no types specified', (done) => {
    const cmd = new MockCommand4();
    const delimiter = (vorpal as any)._delimiter;
    const argv = process.argv;
    vorpal.delimiter('');
    sinon.stub(cmd as any, 'initAction').callsFake((args) => {
      try {
        assert.strictEqual(args.options.option1, '00123');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        vorpal.delimiter(delimiter);
        process.argv = argv;
      }
    });
    sinon.stub(process, 'exit');
    cmd.init(vorpal);
    process.argv = ['node', 'm365', 'mock-command', '--option1', '00123'];
    vorpal.parse(['node', 'm365', 'mock-command', '--option1', '00123']);
  });

  it('removes leading zeroes from known options when the command doesn\'t support unknown options', (done) => {
    const cmd = new MockCommand3();
    const delimiter = (vorpal as any)._delimiter;
    const argv = process.argv;
    vorpal.delimiter('');
    sinon.stub(cmd as any, 'initAction').callsFake((args) => {
      try {
        assert.strictEqual(args.options.option1, 123);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        vorpal.delimiter(delimiter);
        process.argv = argv;
      }
    });
    sinon.stub(process, 'exit');
    cmd.init(vorpal);
    process.argv = ['node', 'm365', 'mock-command', '--option1', '00123'];
    vorpal.parse(['node', 'm365', 'mock-command', '--option1', '00123']);
  });

  it('correctly handles error when instance of error returned from the promise', (done) => {
    const cmd = new MockCommand3();
    (cmd as any).handleRejectedODataPromise(new Error('An error has occurred'), undefined, (msg: any): void => {
      try {
        assert.equal(JSON.stringify(msg), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles graph response (code) from the promise', (done) => {
    const errorMessage = "forbidden-message";
    const errorCode = "Access Denied";
    const cmd = new MockCommand3();
    (cmd as any).handleRejectedODataPromise({ error: { error: { message: errorMessage, code: errorCode } } }, undefined, (msg: any): void => {
      try {
        assert.equal(JSON.stringify(msg), JSON.stringify(new CommandError(errorCode + " - " + errorMessage)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles graph response error (without code) from the promise', (done) => {
    const errorMessage = "forbidden-message";
    const cmd = new MockCommand3();
    (cmd as any).handleRejectedODataPromise({ error: { error: { message: errorMessage } } }, undefined, (msg: any): void => {
      try {
        assert.equal(JSON.stringify(msg), JSON.stringify(new CommandError(errorMessage)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('tracks the use of unknown options in telemetry', () => {
    const command = new MockCommand1();
    const actual = {
      prop1: true
    };
    const expected = JSON.stringify({
      prop1: true,
      // this is expected, because we're not tracking the actual value but rather
      // whether the property is used or not, so the tracked value for an unknown
      // property will be always true
      Prop2: true
    });
    command.trackUnknownOptionsPublic(actual, { Prop2: false });
    assert.equal(JSON.stringify(actual), expected);
  });

  it('adds unknown options to payload', () => {
    const command = new MockCommand1();
    const actual = {
      prop1: true
    };
    const expected = JSON.stringify({
      prop1: true,
      Prop2: false
    });
    command.addUnknownOptionsToPayloadPublic(actual, { Prop2: false })
    assert.equal(JSON.stringify(actual), expected);
  });
});