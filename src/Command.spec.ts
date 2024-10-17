import assert from 'assert';
import chalk from 'chalk';
import sinon from 'sinon';
import { z } from 'zod';
import auth from './Auth.js';
import Command, {
  CommandError,
  globalOptionsZod
} from './Command.js';
import { CommandOptionInfo } from './cli/CommandOptionInfo.js';
import { Logger } from './cli/Logger.js';
import { cli } from './cli/cli.js';
import { telemetry } from './telemetry.js';
import { accessToken } from './utils/accessToken.js';
import { pid } from './utils/pid.js';
import { session } from './utils/session.js';
import { sinonUtil } from './utils/sinonUtil.js';

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
    this.validators.push(async () => { return true; });
  }

  public async commandAction(logger: Logger): Promise<void> {
    await this.showDeprecationWarning(logger, 'mc1', this.name);
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

class MockCommandWithSchema extends Command {
  public get name(): string {
    return 'mock-command';
  }

  public get description(): string {
    return 'Mock command description';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return globalOptionsZod;
  }

  public optionsInfo: CommandOptionInfo[] = [
    {
      name: 'requiredOption',
      required: true,
      type: 'string'
    },
    {
      name: 'optionalString',
      required: false,
      type: 'string'
    },
    {
      name: 'optionalEnum',
      required: false,
      type: 'string',
      autocomplete: ['a', 'b', 'c']
    },
    {
      name: 'optionalNumber',
      required: false,
      type: 'number'
    },
    {
      name: 'optionalBoolean',
      required: false,
      type: 'boolean'
    }
  ];

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  public async commandAction(logger: Logger, args: any): Promise<void> {
    throw 'Exception';
  }
}

describe('Command', () => {
  let telemetryCommandName: any;
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').callsFake((commandName) => {
      telemetryCommandName = commandName;
    });
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  beforeEach(() => {
    telemetryCommandName = null;
    auth.connection.active = true;
    cli.currentCommandName = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      process.exit,
      accessToken.isAppOnlyAccessToken,
      accessToken.getUserIdFromAccessToken
    ]);
    auth.connection.active = false;
  });

  after(() => {
    sinon.restore();
    auth.connection.accessTokens = {};
  });

  it('returns true by default', async () => {
    const cmd = new MockCommand2();
    assert.strictEqual(await cmd.validate({ options: {} }, cli.getCommandInfo(cmd)), true);
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
      assert(loggerLogToStderrSpy.calledWith(chalk.yellow(`Command 'mc1' is deprecated. Please use 'mock-command' instead.`)));
    }
    catch (err: any) {
      assert.fail(err);
    }
  });

  it('logs command name in the telemetry when command name used', async () => {
    const mock = new MockCommand1();
    await mock.action(logger, { options: {} });

    assert.strictEqual(telemetryCommandName, 'mock-command');
  });

  it('logs command alias in the telemetry when command alias used', async () => {
    cli.currentCommandName = 'mc1';
    const mock = new MockCommand1();
    await mock.action(logger, { options: {} });

    assert.strictEqual(telemetryCommandName, 'mc1');
  });

  it('logs empty command name in telemetry when command called using something else than name or alias', async () => {
    cli.currentCommandName = 'foo';
    const mock = new MockCommand1();
    await mock.action(logger, { options: {} });

    assert.strictEqual(telemetryCommandName, '');
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

  it('correctly tracks properties based on schema', () => {
    const command = new MockCommandWithSchema();
    const args = {
      options: {
        requiredOption: 'abc',
        optionalString: 'def',
        optionalEnum: 'a',
        optionalNumber: 1,
        optionalBoolean: false
      }
    };
    const telemetryProps = (command as any).getTelemetryProperties(args);
    assert.deepStrictEqual(telemetryProps, {
      optionalString: true,
      optionalEnum: 'a',
      optionalNumber: true,
      optionalBoolean: false
    });
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

  it('catches exception thrown by commandAction', async () => {
    const command = new MockCommand4();
    await assert.rejects(command.action(logger, { options: {} }), new CommandError('Exception'));
  });

  it('prints command name as the h1 heading in md output', () => {
    const command = new MockCommand1();
    const commandOutput = [{}];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`# mock-command`) > -1);
  });

  it('uses the title property as the preferred item heading', () => {
    const command = new MockCommand1();
    const commandOutput = [{ title: 'title', Title: 'Title', displayName: 'displayName', DisplayName: 'DisplayName', name: 'name', Name: 'Name' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## title`) > -1);
  });

  it('uses the Title property as the item heading when title not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ Title: 'Title', displayName: 'displayName', DisplayName: 'DisplayName', name: 'name', Name: 'Name' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## Title`) > -1);
  });

  it('uses the displayName property as the item heading when title and Title not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ displayName: 'displayName', DisplayName: 'DisplayName', name: 'name', Name: 'Name' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## displayName`) > -1);
  });

  it('uses the DisplayName property as the item heading when title, Title, and displayName not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ DisplayName: 'DisplayName', name: 'name', Name: 'Name' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## DisplayName`) > -1);
  });

  it('uses the name property as the item heading when title, Title, displayName, and DisplayName not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ name: 'name', Name: 'Name' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## name`) > -1);
  });

  it('uses the Name property as the item heading when title, Title, displayName, DisplayName, and name not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ Name: 'Name' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## Name`) > -1);
  });

  it('uses the id property as the preferred item ID', () => {
    const command = new MockCommand1();
    const commandOutput = [{ id: 'id', Id: 'Id', ID: 'ID', uniqueId: 'uniqueId', UniqueId: 'UniqueId', objectId: 'objectId', ObjectId: 'ObjectId', url: 'url', Url: 'Url', URL: 'URL' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## id`) > -1);
  });

  it('uses the Id property as the item ID when id not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ Id: 'Id', ID: 'ID', uniqueId: 'uniqueId', UniqueId: 'UniqueId', objectId: 'objectId', ObjectId: 'ObjectId', url: 'url', Url: 'Url', URL: 'URL' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## Id`) > -1);
  });

  it('uses the ID property as the item ID when id, and Id not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ ID: 'ID', uniqueId: 'uniqueId', UniqueId: 'UniqueId', objectId: 'objectId', ObjectId: 'ObjectId', url: 'url', Url: 'Url', URL: 'URL' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## ID`) > -1);
  });

  it('uses the uniqueId property as the item ID when id, Id, and ID not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ uniqueId: 'uniqueId', UniqueId: 'UniqueId', objectId: 'objectId', ObjectId: 'ObjectId', url: 'url', Url: 'Url', URL: 'URL' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## uniqueId`) > -1);
  });

  it('uses the UniqueId property as the item ID when id, Id, ID, and uniqueId not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ UniqueId: 'UniqueId', objectId: 'objectId', ObjectId: 'ObjectId', url: 'url', Url: 'Url', URL: 'URL' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## UniqueId`) > -1);
  });

  it('uses the objectId property as the item ID when id, Id, ID, uniqueId, and UniqueId not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ objectId: 'objectId', ObjectId: 'ObjectId', url: 'url', Url: 'Url', URL: 'URL' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## objectId`) > -1);
  });

  it('uses the ObjectId property as the item ID when id, Id, ID, uniqueId, UniqueId, and objectId not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ ObjectId: 'ObjectId', url: 'url', Url: 'Url', URL: 'URL' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## ObjectId`) > -1);
  });

  it('uses the url property as the item ID when id, Id, ID, uniqueId, UniqueId, objectId, and ObjectId not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ url: 'url', Url: 'Url', URL: 'URL' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## url`) > -1);
  });

  it('uses the Url property as the item ID when id, Id, ID, uniqueId, UniqueId, objectId, ObjectId, and url not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ Url: 'Url', URL: 'URL' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## Url`) > -1);
  });

  it('uses the URL property as the item ID when id, Id, ID, uniqueId, UniqueId, objectId, ObjectId, url, and URL not set', () => {
    const command = new MockCommand1();
    const commandOutput = [{ URL: 'URL' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## URL`) > -1);
  });

  it('uses the Title property as the preferred item title and the id property as the preferred item ID in the heading', () => {
    const command = new MockCommand1();
    const commandOutput = [{ Title: 'Title', id: 'id' }];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`## Title (id)`) > -1);
  });

  it('properly handles logging no output', () => {
    const command = new MockCommand1();
    const actual = command.getMdOutput(undefined as any, command, { options: { output: 'md' } });
    assert(actual.indexOf(`# mock-command`) > -1);
  });

  it('properly handles logging empty output', () => {
    const command = new MockCommand1();
    const commandOutput: any[] = [];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(`# mock-command`) > -1);
  });

  it('properly handles logging mixed output with an empty item in between', () => {
    const command = new MockCommand1();
    const commandOutput = [
      {
        id: 'id1'
      },
      undefined,
      {
        id: 'id2'
      }
    ];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    const match = actual.match(/## id/g);
    assert.strictEqual(match?.length, 2);
  });

  it('escapes reserved md characters in property names', () => {
    const command = new MockCommand1();
    const commandOutput = [
      {
        '_*~`|': 'value'
      }
    ];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf('\\_\\*\\~\\`\\|') > -1);
  });

  it('escapes reserved md characters in property values', () => {
    const command = new MockCommand1();
    const commandOutput = [
      {
        'property': '_*~`|'
      }
    ];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf('\\_\\*\\~\\`\\|') > -1);
  });

  it('excludes objects in md output', () => {
    const command = new MockCommand1();
    const commandOutput = [
      {
        'property': {
          'property': 'value'
        }
      }
    ];
    const actual = command.getMdOutput(commandOutput, command, { options: { output: 'md' } });
    assert(actual.indexOf(JSON.stringify(commandOutput[0].property)) === -1);
  });

  it('excludes objects that are values to JSON', async () => {
    const command = new MockCommand1();
    const commandOutput = [
      {
        'property': {
          'property': 'value'
        }
      }
    ];
    const actual = await command.getCsvOutput(commandOutput, { options: { output: 'csv' } });
    assert(actual.indexOf(JSON.stringify(commandOutput[0].property)) === -1);
  });

  it('correctly serialize bool values to csv output', async () => {
    const command = new MockCommand1();
    const commandOutput = [
      {
        'property1': true,
        'property2': false
      }
    ];
    const actual = await command.getCsvOutput(commandOutput, { options: { output: 'csv' } });
    assert.strictEqual(actual,"property1,property2\n1,0\n");
  });

  it('passes validation when csv output specified', async () => {
    const cmd = new MockCommand2();
    assert.strictEqual(await cmd.validate({ options: { output: 'csv' } }, cli.getCommandInfo(cmd)), true);
  });

  it('passes validation when json output specified', async () => {
    const cmd = new MockCommand2();
    assert.strictEqual(await cmd.validate({ options: { output: 'json' } }, cli.getCommandInfo(cmd)), true);
  });

  it('passes validation when text output specified', async () => {
    const cmd = new MockCommand2();
    assert.strictEqual(await cmd.validate({ options: { output: 'text' } }, cli.getCommandInfo(cmd)), true);
  });

  it('passes validation when no output specified', async () => {
    const cmd = new MockCommand2();
    assert.strictEqual(await cmd.validate({ options: {} }, cli.getCommandInfo(cmd)), true);
  });

  it('fails validation when invalid output specified', async () => {
    const cmd = new MockCommand2();
    assert.notStrictEqual(await cmd.validate({ options: { output: 'invalid' } }, cli.getCommandInfo(cmd)), true);
  });

  it('handles option with @meid token and spaces', async () => {
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: '',
      accessToken: ''
    };
    const command = new MockCommand3();
    const commandCommandActionSpy: sinon.SinonSpy = sinon.spy(command, 'commandAction');
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns('f3e59491-fc1a-47cc-a1f0-95ed45983717');

    await command.action(logger, { options: { option1: '@Meid ' } });
    assert.deepStrictEqual(commandCommandActionSpy.lastCall.args[1].options.option1, 'f3e59491-fc1a-47cc-a1f0-95ed45983717');
  });

  it('handles option with @meusername token and spaces', async () => {
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: '',
      accessToken: ''
    };
    const command = new MockCommand3();
    const commandCommandActionSpy: sinon.SinonSpy = sinon.spy(command, 'commandAction');
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns('admin@contoso.onmicrosoft.com');

    await command.action(logger, { options: { option1: '@MeUsername ' } });
    assert.deepStrictEqual(commandCommandActionSpy.lastCall.args[1].options.option1, 'admin@contoso.onmicrosoft.com');
  });

  it('handles @meid with application permissions', async () => {
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: '',
      accessToken: ''
    };
    const command = new MockCommand3();
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { option1: '@meId' } }), new CommandError(`It's not possible to use @meId with application permissions`));
  });

  it('handles @meusername with application permissions', async () => {
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: '',
      accessToken: ''
    };
    const command = new MockCommand3();
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { option1: '@meUsername' } }), new CommandError(`It's not possible to use @meUsername with application permissions`));
  });
});