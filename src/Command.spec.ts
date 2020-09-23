import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from './appInsights';
import auth from './Auth';
import { Cli } from './cli';
import { Logger } from './cli/Logger';
import Command, {
  CommandError, CommandOption,
  CommandTypes
} from './Command';
import Utils from './Utils';

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

  public commandAction(logger: Logger, args: any, cb: (err?: any) => void): void {
    this.showDeprecationWarning(logger, 'mc1', this.name);

    cb();
  }

  public validate(): boolean | string {
    return true;
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

  public handlePromiseError(response: any, logger: Logger, callback: (err?: any) => void): void {
    this.handleRejectedODataJsonPromise(response, logger, callback);
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

describe('Command', () => {
  let telemetry: any;
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;
  let cli: Cli;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
    logger = {
      log: (msg: string) => { }
    };
    loggerSpy = sinon.spy(logger, 'log');
    cli = Cli.getInstance();
  });

  beforeEach(() => {
    telemetry = null;
    auth.service.connected = true;
    cli.currentCommandName = undefined;
  });

  afterEach(() => {
    Utils.restore([
      process.exit
    ]);
    auth.service.connected = false;
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
  });

  it('has no autocomplete by default', () => {
    const cmd = new MockCommand2();
    assert.strictEqual(typeof cmd.autocomplete(), 'undefined');
  });

  it('returns true by default', () => {
    const cmd = new MockCommand2();
    assert.strictEqual(cmd.validate({}), true);
  });

  it('does not define option types by default', () => {
    const cmd = new MockCommand2();
    assert.strictEqual(typeof cmd.types(), 'undefined');
  });

  it('removes optional arguments from command name', () => {
    const cmd = new MockCommand2();
    assert.strictEqual(cmd.getCommandName(), 'Mock command 2');
  });

  it('displays error message when it\'s serialized in the error property', () => {
    const mock = new MockCommand2();
    mock.handlePromiseError({
      error: JSON.stringify({
        error: {
          message: 'An error has occurred'
        }
      })
    }, logger, (err?: any) => {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
    });
  });

  it('displays the raw error message when the serialized value from the error property is not an error object', () => {
    const mock = new MockCommand2();
    mock.handlePromiseError({
      error: JSON.stringify({
        error: {
          id: '123'
        }
      })
    }, logger, (err?: any) => {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(JSON.stringify({
        error: {
          id: '123'
        }
      }))));
    });
  });

  it('displays the raw error message when the serialized value from the error property is not a JSON object', () => {
    const mock = new MockCommand2();
    mock.handlePromiseError({
      error: 'abc'
    }, logger, (err?: any) => {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
    });
  });

  it('displays error message coming from ADALJS', () => {
    const mock = new MockCommand2();
    mock.handlePromiseError({
      error: { error_description: 'abc' }
    }, logger, (err?: any) => {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
    });
  });

  it('shows deprecation warning when command executed using the deprecated name', () => {
    cli.currentCommandName = 'mc1';
    const mock = new MockCommand1();
    mock.commandAction(logger, {}, (err?: any): void => {
      assert(loggerSpy.calledWith(chalk.yellow(`Command 'mc1' is deprecated. Please use 'mock-command' instead`)))
    });
  });

  it('logs command name in the telemetry when command name used', (done) => {
    const mock = new MockCommand1();
    mock.action(logger, { options: {} }, () => {
      try {
        assert.strictEqual(telemetry.name, 'mock-command');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs command alias in the telemetry when command alias used', (done) => {
    cli.currentCommandName = 'mc1';
    const mock = new MockCommand1();
    mock.action(logger, { options: {} }, () => {
      try {
        assert.strictEqual(telemetry.name, 'mc1');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs empty command name in telemetry when command called using something else than name or alias', (done) => {
    cli.currentCommandName = 'foo';
    const mock = new MockCommand1();
    mock.action(logger, { options: {} }, () => {
      try {
        assert.strictEqual(telemetry.name, '');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when instance of error returned from the promise', (done) => {
    const cmd = new MockCommand3();
    (cmd as any).handleRejectedODataPromise(new Error('An error has occurred'), undefined, (msg: any): void => {
      try {
        assert.strictEqual(JSON.stringify(msg), JSON.stringify(new CommandError('An error has occurred')));
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
        assert.strictEqual(JSON.stringify(msg), JSON.stringify(new CommandError(errorCode + " - " + errorMessage)));
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
        assert.strictEqual(JSON.stringify(msg), JSON.stringify(new CommandError(errorMessage)));
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
    command.addUnknownOptionsToPayloadPublic(actual, { Prop2: false })
    assert.strictEqual(JSON.stringify(actual), expected);
  });
});