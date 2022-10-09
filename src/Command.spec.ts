import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from './appInsights';
import auth from './Auth';
import { Cli } from './cli/Cli';
import { Logger } from './cli/Logger';
import Command, {
  CommandError
} from './Command';
import { pid } from './utils/pid';
import { sinonUtil } from './utils/sinonUtil';

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

  public allowUnknownOptions(): boolean {
    return true;
  }

  constructor() {
    super();
    
    this.types.string.push('option2');
    this.options.push(
      {
        option: '--debug'
      },
      {
        option: '--option1 [option1]'
      },
      {
        option: '--option2 [option2]'
      }
    );
    this.validators.push(() => Promise.resolve(true));
  }

  public async commandAction(logger: Logger): Promise<void> {
    this.showDeprecationWarning(logger, 'mc1', this.name);
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

  public async commandAction(): Promise<void> {
  }

  public commandHelp(args: any, log: (message: string) => void): void {
    log('MockCommand2 help');
  }

  public handlePromiseError(response: any): void {
    this.handleRejectedODataJsonPromise(response);
  }
}

class MockCommand3 extends Command {
  public get name(): string {
    return 'mock-command';
  }

  public get description(): string {
    return 'Mock command description';
  }

  constructor() {
    super();

    this.options.push(
      {
        option: '--debug'
      },
      {
        option: '--option1 [option1]'
      }
    );
  }

  public async commandAction(): Promise<void> {
  }

  public commandHelp(): void {
  }
}

class MockCommand4 extends Command {
  public get name(): string {
    return 'mock-command';
  }

  public get description(): string {
    return 'Mock command description';
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  public async commandAction(logger: Logger, args: any): Promise<void> {
    throw 'Exception';
  }
}

describe('Command', () => {
  let telemetry: any;
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let cli: Cli;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
    logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    cli = Cli.getInstance();
  });

  beforeEach(() => {
    telemetry = null;
    auth.service.connected = true;
    cli.currentCommandName = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      process.exit
    ]);
    auth.service.connected = false;
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      pid.getProcessName,
      auth.restoreAuth
    ]);
  });

  it('returns true by default', async () => {
    const cmd = new MockCommand2();
    assert.strictEqual(await cmd.validate({ options: {} }, Cli.getCommandInfo(cmd)), true);
  });

  it('removes optional arguments from command name', () => {
    const cmd = new MockCommand2();
    assert.strictEqual(cmd.getCommandName(), 'Mock command 2');
  });

  it('returns alias when command ran using an alias', () => {
    const cmd = new MockCommand1();
    assert.strictEqual(cmd.getCommandName('mc1'), 'mc1');
  });

  it('displays error message when it\'s serialized in the error property', () => {
    const mock = new MockCommand2();
    assert.throws(() => mock.handlePromiseError({
      error: JSON.stringify({
        error: {
          message: 'An error has occurred'
        }
      })
    }), new CommandError('An error has occurred'));
  });

  it('displays the raw error message when the serialized value from the error property is not an error object', () => {
    try {
      const mock = new MockCommand2();
      mock.handlePromiseError({
        error: JSON.stringify({
          error: {
            id: '123'
          }
        })
      });
      assert.fail('No exception was thrown.');
    }
    catch (err: any) {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(JSON.stringify({
        error: {
          id: '123'
        }
      }))));
    }
  });

  it('displays the raw error message when the serialized value from the error property is not a JSON object', () => {
    try {
      const mock = new MockCommand2();
      mock.handlePromiseError({
        error: 'abc'
      });
      assert.fail('No exception was thrown.');
    }
    catch (err: any) {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
    }
  });

  it('displays error message coming from ADALJS', () => {
    try {
      const mock = new MockCommand2();
      mock.handlePromiseError({
        error: { error_description: 'abc' }
      });
      assert.fail('No exception was thrown.');
    }
    catch (err: any) {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
    }
  });

  it('shows deprecation warning when command executed using the deprecated name', async () => {
    try {
      cli.currentCommandName = 'mc1';
      const mock = new MockCommand1();
      await mock.commandAction(logger);
      assert(loggerLogToStderrSpy.calledWith(chalk.yellow(`Command 'mc1' is deprecated. Please use 'mock-command' instead`)));
    }
    catch (err: any) {
      assert.fail(err);
    }
  });

  it('logs command name in the telemetry when command name used', async() => {
    const mock = new MockCommand1();
    await mock.action(logger, { options: {} });

    assert.strictEqual(telemetry.name, 'mock-command');
  });

  it('logs command alias in the telemetry when command alias used', async () => {
    cli.currentCommandName = 'mc1';
    const mock = new MockCommand1();
    await mock.action(logger, { options: {} });

    assert.strictEqual(telemetry.name, 'mc1');
  });

  it('logs empty command name in telemetry when command called using something else than name or alias', async () => {
    cli.currentCommandName = 'foo';
    const mock = new MockCommand1();
    await mock.action(logger, { options: {} });

    assert.strictEqual(telemetry.name, '');
  });

  it('correctly handles error when instance of error returned from the promise', () => {
    const cmd = new MockCommand3();
    assert.throws(() => (cmd as any).handleRejectedODataPromise(new Error('An error has occurred')), new CommandError('An error has occurred'));
  });

  it('correctly handles graph response (code) from the promise', () => {
    const errorMessage = "forbidden-message";
    const errorCode = "Access Denied";
    const cmd = new MockCommand3();
    assert.throws(() => (cmd as any).handleRejectedODataPromise({ error: { error: { message: errorMessage, code: errorCode } } }),
      new CommandError(errorCode + " - " + errorMessage));
  });

  it('correctly handles graph response error (without code) from the promise', () => {
    const errorMessage = "forbidden-message";
    const cmd = new MockCommand3();
    assert.throws(() => (cmd as any).handleRejectedODataPromise({ error: { error: { message: errorMessage } } }), new CommandError(errorMessage));
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
    assert.strictEqual(JSON.stringify(actual), expected);
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
    command.addUnknownOptionsToPayloadPublic(actual, { Prop2: false });
    assert.strictEqual(JSON.stringify(actual), expected);
  });

  it('catches exception thrown by commandAction', async() => {
    const command = new MockCommand4();
    await assert.rejects(command.action(logger, { options: {} }), new CommandError('Exception'));
  });
});