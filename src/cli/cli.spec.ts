import assert from 'assert';
import chalk from 'chalk';
import Table from 'easy-table';
import fs from 'fs';
import { createRequire } from 'module';
import os from 'os';
import path from 'path';
import sinon from 'sinon';
import url from 'url';
import Command, { CommandError } from '../Command.js';
import AnonymousCommand from '../m365/base/AnonymousCommand.js';
import cliCompletionUpdateCommand from '../m365/cli/commands/completion/completion-clink-update.js';
import { settingsNames } from '../settingsNames.js';
import { telemetry } from '../telemetry.js';
import { md } from '../utils/md.js';
import { pid } from '../utils/pid.js';
import { Choice, SelectionConfig, prompt } from '../utils/prompt.js';
import { session } from '../utils/session.js';
import { sinonUtil } from '../utils/sinonUtil.js';
import { cli, CommandOutput } from './cli.js';
import { Logger } from './Logger.js';

const require = createRequire(import.meta.url);
const packageJSON = require('../../package.json');
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

class MockCommand extends AnonymousCommand {
  public get name(): string {
    return 'cli mock';
  }
  public get description(): string {
    return 'Mock command';
  }
  constructor() {
    super();

    this.options.push(
      {
        option: '-x, --parameterX <parameterX>'
      },
      {
        option: '-y, --parameterY [parameterY]'
      }
    );
    this.types.string.push('x', 'y');
  }
  public async commandAction(logger: Logger, args: any): Promise<void> {
    await logger.log(args.options.parameterX);
  }
}

class MockCommandWithAutocomplete extends AnonymousCommand {
  public get name(): string {
    return 'cli mock autocomplete';
  }
  public get description(): string {
    return 'Mock command';
  }
  constructor() {
    super();

    this.options.push(
      {
        option: '-x, --parameterX <parameterX>',
        autocomplete: ['value1', 'value2']
      }
    );
    this.types.string.push('x', 'y');
  }
  public async commandAction(logger: Logger, args: any): Promise<void> {
    await logger.log(args.options.parameterX);
  }
}

class MockCommandWithOptionSets extends AnonymousCommand {
  public get name(): string {
    return 'cli mock optionsets';
  }
  public get description(): string {
    return 'Mock command with option sets';
  }
  constructor() {
    super();

    this.options.push(
      {
        option: '--opt1 [name]'
      },
      {
        option: '--opt2 [name]'
      },
      {
        option: '--opt3 [name]'
      },
      {
        option: '--opt4 [name]'
      },
      {
        option: '--opt5 [name]'
      },
      {
        option: '--opt6 [name]'
      }
    );
    this.optionSets.push(
      { options: ['opt1', 'opt2'] },
      {
        options: ['opt3', 'opt4'],
        runsWhen: (args) => typeof args.options.opt2 !== 'undefined' // validate when opt2 is set
      },
      {
        options: ['opt5', 'opt6'],
        runsWhen: (args) => { return args.options.opt5 || args.options.opt6; } // validate when opt5 or opt6 is set
      }
    );
  }
  public async commandAction(): Promise<void> {
  }
}

class MockCommandWithAlias extends AnonymousCommand {
  public get name(): string {
    return 'cli mock alias';
  }
  public get description(): string {
    return 'Mock command with alias';
  }
  public alias(): string[] {
    return ['cli mock alt'];
  }
  public async commandAction(): Promise<void> {
  }
}

class MockCommandWithValidation extends AnonymousCommand {
  public get name(): string {
    return 'cli mock1 validation';
  }
  public get description(): string {
    return 'Mock command with validation';
  }
  constructor() {
    super();

    this.options.push(
      {
        option: '-x, --parameterX <parameterX>'
      },
      {
        option: '-y, --parameterY [parameterY]'
      }
    );
  }
  public async commandAction(): Promise<void> {
  }
}

class MockCommandWithBooleanRewrite extends AnonymousCommand {
  public get name(): string {
    return 'cli mock boolean rewrite';
  }
  public get description(): string {
    return 'Mock command with boolean rewrite';
  }
  constructor() {
    super();

    this.options.push(
      {
        option: '-x, --booleanParameterX [booleanParameterX]'
      },
      {
        option: '-y, --booleanParameterY [booleanParameterY]'
      }
    );

    this.types.boolean.push('x', 'booleanParameterX', 'y', 'booleanParameterY');
  }
  public async commandAction(logger: Logger, args: any): Promise<void> {
    await logger.log(`booleanParameterX: ${args.options.booleanParameterX}`);
    await logger.log(`booleanParameterY: ${args.options.booleanParameterY}`);
  }
}

class MockCommandWithConfirmationPrompt extends AnonymousCommand {
  public get name(): string {
    return 'cli mock prompt';
  }
  public get description(): string {
    return 'Mock command with prompt';
  }
  public async commandAction(): Promise<void> {
    await cli.promptForConfirmation({ message: `Continue?` });
  }
}

class MockCommandWithHandleMultipleResultsFound extends AnonymousCommand {
  public get name(): string {
    return 'cli mock interactive prompt';
  }
  public get description(): string {
    return 'Mock command with interactive prompt';
  }
  public async commandAction(): Promise<void> {
    await cli.handleMultipleResultsFound(`Multiple values with name found.`, { '1': { 'id': '1', 'title': 'Option1' }, '2': { 'id': '2', 'title': 'Option2' } });
  }
}

class MockCommandWithOutput extends AnonymousCommand {
  public get name(): string {
    return 'cli mock output';
  }
  public get description(): string {
    return 'Mock command with output';
  }
  public async commandAction(logger: Logger): Promise<void> {
    await logger.log('Command output');
  }
}

class MockCommandWithRawOutput extends AnonymousCommand {
  public get name(): string {
    return 'cli mock output';
  }
  public get description(): string {
    return 'Mock command with output';
  }
  public async commandAction(logger: Logger): Promise<void> {
    if (this.debug) {
      await logger.logToStderr('Debug output');
    }

    await logger.logRaw('Raw output');
  }
}

describe('cli', () => {
  let rootFolder: string;
  let cliLogStub: sinon.SinonStub;
  let cliErrorStub: sinon.SinonStub;
  let cliFormatOutputSpy: sinon.SinonSpy;
  let processExitStub: sinon.SinonStub;
  let md2plainSpy: sinon.SinonSpy;
  let mockCommandActionSpy: sinon.SinonSpy;
  let mockCommand: Command;
  let mockCommandWithAutocomplete: Command;
  let mockCommandWithOptionSets: Command;
  let mockCommandWithAlias: Command;
  let mockCommandWithValidation: Command;
  let log: string[] = [];
  let mockCommandWithBooleanRewrite: Command;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');

    cliLogStub = sinon.stub(cli, 'log').callsFake(message => {
      log.push(message as string ?? '');
    });
    cliErrorStub = sinon.stub(cli, 'error');
    cliFormatOutputSpy = sinon.spy(cli, 'formatOutput');
    processExitStub = sinon.stub(process, 'exit').callsFake((() => { }) as any);
    md2plainSpy = sinon.spy(md, 'md2plain');

    mockCommand = new MockCommand();
    mockCommandWithAutocomplete = new MockCommandWithAutocomplete();
    mockCommandWithAlias = new MockCommandWithAlias();
    mockCommandWithBooleanRewrite = new MockCommandWithBooleanRewrite();
    mockCommandWithValidation = new MockCommandWithValidation();
    mockCommandWithOptionSets = new MockCommandWithOptionSets();
    mockCommandActionSpy = sinon.spy(mockCommand, 'action');

    return new Promise((resolve) => {
      fs.realpath(__dirname, (err: NodeJS.ErrnoException | null, resolvedPath: string): void => {
        rootFolder = resolvedPath;
        resolve(undefined);
      });
    });
  });

  beforeEach(() => {
    log = [];
    cli.commands = [
      cli.getCommandInfo(mockCommand, 'cli-mock.js', 'help.mdx'),
      cli.getCommandInfo(mockCommandWithAutocomplete, 'cli-autocomplete-mock.js', 'help.mdx'),
      cli.getCommandInfo(mockCommandWithOptionSets, 'cli-optionsets-mock.js', 'help.mdx'),
      cli.getCommandInfo(mockCommandWithAlias, 'cli-alias-mock.js', 'help.mdx'),
      cli.getCommandInfo(mockCommandWithValidation, 'cli-validation-mock.js', 'help.mdx'),
      cli.getCommandInfo(cliCompletionUpdateCommand, 'cli/commands/completion/completion-clink-update.js', 'cli/completion/completion-clink-update.mdx'),
      cli.getCommandInfo(mockCommandWithBooleanRewrite, 'cli-boolean-rewrite-mock.js', 'help.mdx')
    ];
    sinon.stub(cli, 'loadAllCommandsInfo').callsFake(() => '');
    cli.commandToExecute = undefined;
  });

  afterEach(() => {
    cliLogStub.resetHistory();
    cliErrorStub.resetHistory();
    cliFormatOutputSpy.resetHistory();
    processExitStub.reset();
    md2plainSpy.resetHistory();
    mockCommandActionSpy.resetHistory();
    sinonUtil.restore([
      cli.executeCommand,
      fs.existsSync,
      fs.readFileSync,
      // eslint-disable-next-line no-console
      console.log,
      // eslint-disable-next-line no-console
      console.error,
      mockCommand.validate,
      mockCommandWithAutocomplete.validate,
      mockCommandWithValidation.action,
      mockCommandWithValidation.validate,
      mockCommand.commandAction,
      mockCommand.processOptions,
      prompt.forInput,
      prompt.forSelection,
      prompt.forConfirmation,
      cli.getSettingWithDefaultValue,
      cli.loadAllCommandsInfo,
      cli.getConfig().get,
      cli.loadCommandFromFile
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('shows generic help when no command specified', async () => {
    await cli.execute([]);
    assert(cliLogStub.calledWith(`CLI for Microsoft 365 v${packageJSON.version}`));
  });

  it('exits with 0 code when no command specified', async () => {
    await cli.execute([]);
    assert(processExitStub.calledWith(0));
  });

  it('shows generic help when help command and no command name specified', async () => {
    await cli.execute(['help']);
    assert(cliLogStub.calledWith(`CLI for Microsoft 365 v${packageJSON.version}`));
  });

  it('shows generic help when --help option specified', async () => {
    await cli.execute(['--help']);
    assert(cliLogStub.calledWith(`CLI for Microsoft 365 v${packageJSON.version}`));
  });

  it('shows generic help when -h option specified', async () => {
    await cli.execute(['-h']);
    assert(cliLogStub.calledWith(`CLI for Microsoft 365 v${packageJSON.version}`));
  });

  it('shows help for the specific command when help specified followed by a valid command name', async () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().endsWith('.mdx') || path.toString().endsWith('-mock.js'));
    const originalFsReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').returns(originalFsReadFileSync(path.join(rootFolder, '..', '..', 'docs', 'docs', 'cmd', 'cli', 'completion', 'completion-clink-update.mdx'), 'utf8'));
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['help', 'cli', 'mock']);
    assert(md2plainSpy.called);
  });

  it('shows help for the specific command when valid command name specified followed by --help', async () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().endsWith('.mdx') || path.toString().endsWith('-mock.js'));
    const originalFsReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').returns(originalFsReadFileSync(path.join(rootFolder, '..', '..', 'docs', 'docs', 'cmd', 'cli', 'completion', 'completion-clink-update.mdx'), 'utf8'));
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '--help']);
    assert(md2plainSpy.called);
  });

  it('shows help for the specific command when valid command name specified followed by -h', async () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().endsWith('.mdx') || path.toString().endsWith('-mock.js'));
    const originalFsReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').returns(originalFsReadFileSync(path.join(rootFolder, '..', '..', 'docs', 'docs', 'cmd', 'cli', 'completion', 'completion-clink-update.mdx'), 'utf8'));
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '-h']);
    assert(md2plainSpy.called);
  });

  it('shows help for the specific command when valid command name specified followed by -h (single-word command)', async () => {
    sinonUtil.restore(cli.loadAllCommandsInfo);

    await cli.execute(['status', '-h']);
    assert(md2plainSpy.called);
  });

  it('shows help for the specific command when help specified followed by a valid command alias', async () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().endsWith('.mdx') || path.toString().endsWith('-mock.js'));
    const originalFsReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').returns(originalFsReadFileSync(path.join(rootFolder, '..', '..', 'docs', 'docs', 'cmd', 'cli', 'completion', 'completion-clink-update.mdx'), 'utf8'));
    cli.commandToExecute = cli.commands.find(c => c.aliases?.some(a => a === 'cli mock alt'));

    await cli.execute(['help', 'cli', 'mock', 'alt']);
    assert(md2plainSpy.called);
  });

  it('shows full help when specified -h with a number', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').returns('full');

    await cli.execute(['cli', 'completion', 'clink', 'update', '-h', '1']);
    assert(log.some(l => l.indexOf('OPTIONS') > -1), 'Options section not found');
    assert(log.some(l => l.indexOf('EXAMPLES') > -1), 'Examples section not found');
  });

  it('shows full help when specified -h with full', async () => {
    await cli.execute(['cli', 'completion', 'clink', 'update', '-h', 'full']);
    assert(log.some(l => l.indexOf('OPTIONS') > -1), 'Options section not found');
    assert(log.some(l => l.indexOf('EXAMPLES') > -1), 'Examples section not found');
  });

  it('shows help with options section when specified -h with options', async () => {
    await cli.execute(['cli', 'completion', 'clink', 'update', '-h', 'options']);
    assert(log.some(l => l.indexOf('OPTIONS') > -1), 'Options section not found');
    assert(log.some(l => l.indexOf('EXAMPLES') === -1), 'Examples section found');
  });

  it('shows help with examples section when specified -h with examples', async () => {
    await cli.execute(['cli', 'completion', 'clink', 'update', '-h', 'examples']);
    assert(log.some(l => l.indexOf('OPTIONS') === -1), 'Options section found');
    assert(log.some(l => l.indexOf('EXAMPLES') > -1), 'Examples section not found');
  });

  it('shows help with remarks section when specified -h with remarks', async () => {
    await cli.execute(['cli', 'completion', 'clink', 'update', '-h', 'remarks']);
    assert(log.some(l => l.indexOf('REMARKS') > -1), 'Remarks section not found');
    assert(log.some(l => l.indexOf('OPTIONS') === -1), 'Options section found');
  });

  it('shows error when specified -h with an invalid value', async () => {
    try {
      await cli.execute(['cli', 'completion', 'clink', 'update', '-h', 'invalid']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(cliErrorStub.getCalls().some(c => c.firstArg.indexOf('Unknown help mode invalid. Allowed values are') > -1));
    }

  });

  it(`passes options validation if the command doesn't allow unknown options and specified options match command options`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '-x', '123', '-y', '456']);
    assert(mockCommandActionSpy.called);
  });

  it(`succeeds running with truthy/falsy values 'true' and 'false'`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock boolean rewrite');

    await cli.execute(['cli', 'mock', 'boolean', 'rewrite', '--booleanParameterX', 'true', '--booleanParameterY', 'false', '--output', 'text']);
    assert(cliLogStub.calledWith(`booleanParameterX: true`));
    assert(cliLogStub.calledWith(`booleanParameterY: false`));
  });

  it(`rewrites a truthy/falsy values '1' and '0' to 'true' and 'false' respectively`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock boolean rewrite');

    await cli.execute(['cli', 'mock', 'boolean', 'rewrite', '--booleanParameterX', '1', '--booleanParameterY', '0', '--output', 'text']);
    assert(cliLogStub.calledWith(`booleanParameterX: true`));
    assert(cliLogStub.calledWith(`booleanParameterY: false`));
  });

  it(`rewrites a truthy/falsy values 'on' and 'off' to 'true' and 'false' respectively`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock boolean rewrite');

    await cli.execute(['cli', 'mock', 'boolean', 'rewrite', '--booleanParameterX', 'on', '--booleanParameterY', 'off', '--output', 'text']);
    assert(cliLogStub.calledWith(`booleanParameterX: true`));
    assert(cliLogStub.calledWith(`booleanParameterY: false`));
  });

  it(`rewrites a truthy/falsy values 'yes' and 'no' to 'true' and 'false' respectively`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock boolean rewrite');

    await cli.execute(['cli', 'mock', 'boolean', 'rewrite', '--booleanParameterX', 'yes', '--booleanParameterY', 'no', '--output', 'text']);
    assert(cliLogStub.calledWith(`booleanParameterX: true`));
    assert(cliLogStub.calledWith(`booleanParameterY: false`));
  });

  it(`rewrites a truthy/falsy values 'True' and 'False' to 'true' and 'false' respectively`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock boolean rewrite');

    await cli.execute(['cli', 'mock', 'boolean', 'rewrite', '--booleanParameterX', 'True', '--booleanParameterY', 'False', '--output', 'text']);
    assert(cliLogStub.calledWith(`booleanParameterX: true`));
    assert(cliLogStub.calledWith(`booleanParameterY: false`));
  });

  it(`rewrites a truthy/falsy values 'yes' and 'no' to 'true' and 'false' respectively (using shorts)`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock boolean rewrite');

    await cli.execute(['cli', 'mock', 'boolean', 'rewrite', '-x', 'yes', '-y', 'no', '--output', 'text']);
    assert(cliLogStub.calledWith(`booleanParameterX: true`));
    assert(cliLogStub.calledWith(`booleanParameterY: false`));
  });

  it(`shows error when a boolean option does not contain a correct truthy/falsy value`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock boolean rewrite');

    try {
      await cli.execute(['cli', 'mock', 'boolean', 'rewrite', '--booleanParameterX', 'folse']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(cliErrorStub.calledWith(chalk.red(`Error: The value 'folse' for option '--booleanParameterX' is not a valid boolean`)));
    }
  });

  it(`fails options validation if the command doesn't allow unknown options and specified options match command options`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    try {
      await cli.execute(['cli', 'mock', '-x', '123', '--paramZ']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(cliErrorStub.calledWith(chalk.red(`Error: Invalid option: 'paramZ'${os.EOL}`)));
    }
  });

  it(`doesn't execute command action when option validation failed`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    try {
      await cli.execute(['cli', 'mock', '-x', '123', '--paramZ']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(mockCommandActionSpy.notCalled);
    }
  });

  it(`exits with exit code 1 when option validation failed`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    try {
      await cli.execute(['cli', 'mock', '-x', '123', '--paramZ']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(processExitStub.calledWith(1));
    }

  });

  it(`does not prompt and fails validation if a required option is missing`, async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return undefined;
      }
      return defaultValue;
    });

    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    try {
      await cli.execute(['cli', 'mock']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(cliErrorStub.calledWith(chalk.red(`Error: Required option parameterX not specified`)));
    }
  });

  it(`shows validation error when no option from a required set is specified`, async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');

    try {
      await cli.execute(['cli', 'mock', 'optionsets']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(cliErrorStub.calledWith(chalk.red('Error: Specify one of the following options: opt1, opt2.')));
    }
  });

  it(`shows validation error when multiple options from a required set are specified`, async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');

    try {
      await cli.execute(['cli', 'mock', 'optionsets', '--opt1', 'testvalue', '--opt2', 'testvalue']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(cliErrorStub.calledWith(chalk.red('Error: Specify one of the following options: opt1, opt2, but not multiple.')));
    }
  });

  it(`passes validation when one option from a required set is specified`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');

    await cli.execute(['cli', 'mock', 'optionsets', '--opt1', 'testvalue']);
    assert(cliErrorStub.notCalled);
  });

  it(`shows validation error when no option from a dependent set is set`, async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');
    try {
      await cli.execute(['cli', 'mock', 'optionsets', '--opt2', 'testvalue']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(cliErrorStub.calledWith(chalk.red('Error: Specify one of the following options: opt3, opt4.')));
    }
  });

  it(`passes validation when one option from a dependent set is specified`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');

    await cli.execute(['cli', 'mock', 'optionsets', '--opt2', 'testvalue', '--opt3', 'testvalue']);
    assert(cliErrorStub.notCalled);
  });

  it(`shows validation error when multiple options from an optional set are specified`, async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');

    try {
      await cli.execute(['cli', 'mock', 'optionsets', '--opt1', 'testvalue', '--opt5', 'testvalue', '--opt6', 'testvalue']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(cliErrorStub.calledWith(chalk.red('Error: Specify one of the following options: opt5, opt6, but not multiple.')));
    }
  });

  it(`passes validation when one option from an optional set is specified`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');

    await cli.execute(['cli', 'mock', 'optionsets', '--opt2', 'testvalue', '--opt3', 'testvalue', '--opt5', 'testvalue']);
    assert(cliErrorStub.notCalled);
  });

  it(`prompts for required options`, async () => {
    const promptStub: sinon.SinonStub = sinon.stub(prompt, 'forInput').resolves("test");
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return 'true';
      }
      return defaultValue;
    });

    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock']);
    assert(promptStub.called);
  });

  it(`prompts for required options when autocomplete has items`, async () => {
    const promptStub: sinon.SinonStub = sinon.stub(prompt, 'forSelection').resolves("value1");
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return 'true';
      }
      if (settingName === settingsNames.promptListPageSize) {
        return 10;
      }
      return defaultValue;
    });

    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock autocomplete');
    await cli.execute(['cli', 'mock', 'autocomplete']);
    assert.strictEqual(promptStub.firstCall.args[0].choices[0].value, 'value1');
    assert.strictEqual(promptStub.firstCall.args[0].choices[1].value, 'value2');
    assert(promptStub.calledOnce);
  });

  it(`prompts for optionset name and value when optionset not specified`, async () => {
    let firstOptionValue = '', secondOptionValue = '';
    const promptStub: sinon.SinonStub = sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      firstOptionValue = (config.choices[0] as Choice<any>).value;
      secondOptionValue = (config.choices[1] as Choice<any>).value;
      return (config.choices[0] as Choice<any>).value;
    });

    const promptInputStub: sinon.SinonStub = sinon.stub(prompt, 'forInput').resolves('Test 123');

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return 'true';
      }
      return defaultValue;
    });
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');
    await cli.execute(['cli', 'mock', 'optionsets']);
    assert.strictEqual(promptStub.firstCall.args[0].choices[0].value, firstOptionValue);
    assert.strictEqual(promptStub.firstCall.args[0].choices[1].value, secondOptionValue);
    assert.strictEqual(promptInputStub.firstCall.args[0].message, `${firstOptionValue}:`);
    assert(promptStub.calledOnce);
    assert(promptInputStub.calledOnce);
  });

  it(`prompts to choose which option you wish to use when multiple options in a specific optionset are specified`, async () => {
    let firstOptionValue = '', secondOptionValue = '';
    const promptStub: sinon.SinonStub = sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      firstOptionValue = (config.choices[0] as Choice<any>).value;
      secondOptionValue = (config.choices[1] as Choice<any>).value;
      return (config.choices[0] as Choice<any>).value;
    });

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return 'true';
      }
      return defaultValue;
    });
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');
    await cli.execute(['cli', 'mock', 'optionsets', '--opt1', 'testvalue', '--opt2', 'testvalue']);
    assert.strictEqual(promptStub.lastCall.args[0].message, `Option to use:`);
    assert.strictEqual(promptStub.lastCall.args[0].choices[0].value, firstOptionValue);
    assert.strictEqual(promptStub.lastCall.args[0].choices[1].value, secondOptionValue);
    assert(promptStub.calledOnce);
  });

  it(`prompts to choose runsWhen option from optionSet when dependant option is set and prompts for the value`, async () => {
    let firstOptionValue = '', secondOptionValue = '';
    const promptStub: sinon.SinonStub = sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      firstOptionValue = (config.choices[0] as Choice<any>).value;
      secondOptionValue = (config.choices[1] as Choice<any>).value;
      return (config.choices[0] as Choice<any>).value;
    });

    const promptInputStub: sinon.SinonStub = sinon.stub(prompt, 'forInput').resolves('Test 123');

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return 'true';
      }
      return defaultValue;
    });
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');
    await cli.execute(['cli', 'mock', 'optionsets', '--opt2', 'testvalue']);
    assert.strictEqual(promptStub.firstCall.args[0].message, `Option to use:`);
    assert.strictEqual(promptStub.firstCall.args[0].choices[0].value, firstOptionValue);
    assert.strictEqual(promptStub.firstCall.args[0].choices[1].value, secondOptionValue);
    assert(promptStub.calledOnce);
    assert(promptInputStub.calledOnce);
  });

  it(`prompts to pick one of the options from an optionSet when runsWhen condition is matched`, async () => {
    let firstOptionValue = '', secondOptionValue = '';
    const promptStub: sinon.SinonStub = sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      firstOptionValue = (config.choices[0] as Choice<any>).value;
      secondOptionValue = (config.choices[1] as Choice<any>).value;
      return (config.choices[0] as Choice<any>).value;
    });

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return 'true';
      }
      return defaultValue;
    });
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock optionsets');
    await cli.execute(['cli', 'mock', 'optionsets', '--opt2', 'testvalue', '--opt3', 'opt 3', '--opt4', 'opt 4']);
    assert.strictEqual(promptStub.lastCall.args[0].choices[0].value, firstOptionValue);
    assert.strictEqual(promptStub.lastCall.args[0].choices[1].value, secondOptionValue);
    assert(promptStub.calledOnce);
  });

  it(`calls command's validation method when defined`, async () => {
    const mockCommandValidateSpy: sinon.SinonSpy = sinon.spy(mockCommandWithValidation, 'validate');
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock1 validation');

    await cli.execute(['cli', 'mock1', 'validation', '-x', '123']);
    assert(mockCommandValidateSpy.called);
  });

  it(`passes validation when the command's validate method returns true`, async () => {
    sinon.stub(mockCommandWithValidation, 'validate').resolves(true);
    const mockCommandWithValidationActionSpy: sinon.SinonSpy = sinon.spy(mockCommandWithValidation, 'action');

    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock1 validation');

    await cli.execute(['cli', 'mock1', 'validation', '-x', '123']);
    assert(mockCommandWithValidationActionSpy.called);
  });

  it(`fails validation when the command's validate method returns a string`, async () => {
    sinon.stub(mockCommandWithValidation, 'validate').resolves('Error');
    const mockCommandWithValidationActionSpy: sinon.SinonSpy = sinon.spy(mockCommandWithValidation, 'action');

    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock1 validation');

    try {
      await cli.execute(['cli', 'mock1', 'validation', '-x', '123']);
      assert.fail('Error not thrown');
    }
    catch {
      assert(mockCommandWithValidationActionSpy.notCalled);
    }
  });

  it(`executes command when validation passed`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '-x', '123']);
    assert(mockCommandActionSpy.called);
  });

  it(`writes DONE when executing command in verbose mode succeeded`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '-x', '123', '--verbose']);
    assert(cliErrorStub.calledWith(chalk.green('DONE')));
  });

  it(`writes DONE when executing command in debug mode succeeded`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '-x', '123', '--debug']);
    assert(cliErrorStub.calledWith(chalk.green('DONE')));
  });

  it('executes the specified command', async () => {
    await cli.executeCommand(mockCommand, { options: { _: [] } });
    assert(mockCommandActionSpy.called);
  });

  it('logs command name when executing command in debug mode', async () => {
    await cli.executeCommand(mockCommand, { options: { debug: true, _: [] } });
    assert(cliErrorStub.calledWith('Executing command cli mock with options {"options":{"debug":true,"_":[]}}'));
  });

  it('calls confirmation prompt tool when command shows prompt', async () => {
    const promptStub: sinon.SinonStub = sinon.stub(prompt, 'forConfirmation').resolves(true);
    const mockCommandWithConfirmationPrompt = new MockCommandWithConfirmationPrompt();

    await cli.executeCommand(mockCommandWithConfirmationPrompt, { options: { _: [] } });
    assert(promptStub.called);
  });

  it('prints command output with formatting', async () => {
    const commandWithOutput: MockCommandWithOutput = new MockCommandWithOutput();

    await cli.executeCommand(commandWithOutput, { options: { _: [] } });
    assert(cliLogStub.called, 'cli.log not called');
    assert(cliFormatOutputSpy.called, 'cli.formatOutput not called');
  });

  it('prints command output without formatting', async () => {
    const commandWithOutput: MockCommandWithRawOutput = new MockCommandWithRawOutput();

    await cli.executeCommand(commandWithOutput, { options: { _: [] } });
    assert(cliLogStub.called, 'cli.log not called');
    assert(cliFormatOutputSpy.notCalled, 'cli.formatOutput called');
  });

  it('returns command output when executing command with output', async () => {
    const commandWithOutput: MockCommandWithOutput = new MockCommandWithOutput();

    const output: CommandOutput = await cli.executeCommandWithOutput(commandWithOutput, { options: { _: [], output: 'text' } });
    assert.strictEqual(output.stdout, 'Command output');
  });

  it('returns raw command output when executing command with output', async () => {
    const commandWithOutput: MockCommandWithRawOutput = new MockCommandWithRawOutput();

    const output = await cli.executeCommandWithOutput(commandWithOutput, { options: { _: [], output: 'text' } });
    assert.strictEqual(output.stdout, 'Raw output');
  });

  it('returns debug command output when executing command with output in debug mode', async () => {
    const commandWithOutput: MockCommandWithRawOutput = new MockCommandWithRawOutput();

    const output = await cli.executeCommandWithOutput(commandWithOutput, { options: { _: [], debug: true, output: 'text' } });
    assert.strictEqual(output.stdout, 'Raw output');
    assert.strictEqual(output.stderr, ['Executing command cli mock output with options {"options":{"_":[],"debug":true,"output":"text"}}', 'Debug output'].join(os.EOL));
  });

  it('captures command stdout output in a listener when specified', async () => {
    let output: string = '';
    const commandWithOutput: MockCommandWithOutput = new MockCommandWithOutput();

    await cli.executeCommandWithOutput(commandWithOutput, { options: { _: [], output: 'text' } }, {
      stdout: (message) => output = message
    });
    assert.strictEqual(output, 'Command output');
  });

  it('captures command raw stdout output in a listener when specified', async () => {
    let output: string = '';
    const commandWithOutput: MockCommandWithRawOutput = new MockCommandWithRawOutput();

    await cli.executeCommandWithOutput(commandWithOutput, { options: { _: [], output: 'text' } }, {
      stdout: (message) => output = message
    });
    assert.strictEqual(output, 'Raw output');
  });

  it('captures command stderr output in a listener when specified', async () => {
    const output: string[] = [];
    const commandWithOutput: MockCommandWithRawOutput = new MockCommandWithRawOutput();
    await cli.executeCommandWithOutput(commandWithOutput, { options: { _: [], output: 'text', debug: true } }, {
      stderr: (message) => output.push(message)
    });

    assert.deepStrictEqual(output, ['Executing command cli mock output with options {"options":{"_":[],"output":"text","debug":true}}', 'Debug output']);
  });

  it('calls prompt tool when command shows prompt and executed with output', async () => {
    const promptStub: sinon.SinonStub = sinon.stub(prompt, 'forConfirmation').resolves(true);
    const mockCommandWithConfirmationPrompt = new MockCommandWithConfirmationPrompt();

    await cli.executeCommandWithOutput(mockCommandWithConfirmationPrompt, { options: { _: [] } });
    assert(promptStub.called);
  });

  it('calls prompt tool when command shows interactive prompt and executed with output', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').returns(true);

    const promptStub: sinon.SinonStub = sinon.stub(prompt, 'forSelection').resolves("test" as any);
    const mockCommandWithHandleMultipleResultsFound = new MockCommandWithHandleMultipleResultsFound();

    await cli.executeCommandWithOutput(mockCommandWithHandleMultipleResultsFound, { options: { _: [] } });
    assert(promptStub.called);
  });

  it('throws error when interactive mode not set', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').returns(false);
    await assert.rejects((cli.handleMultipleResultsFound(`Multiple values with name found.`, { '1': { 'id': '1', 'title': 'Option1' }, '2': { 'id': '2', 'title': 'Option2' } })
    ), 'error');
  });

  it('correctly handles error when executing command', async () => {
    sinon.stub(mockCommand, 'commandAction').throws(new Error('Error'));

    try {
      await cli.executeCommand(mockCommand, { options: { _: [] } });
      assert.fail('Command succeeded while expected fail');
    }
    catch (e: any) {
      assert.strictEqual(e.message, 'Error');
    }
  });

  it('correctly handles error when executing command (execute)', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(() => Promise.reject('Error'));
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli completion clink update');
    assert.rejects(cli.execute(['cli', 'completion', 'clink', 'update']), new Error('Error'));
  });

  it('correctly handles error when executing command with output', async () => {
    sinon.stub(mockCommand, 'commandAction').throws(new Error('Error'));

    try {
      await cli.executeCommandWithOutput(mockCommand, { options: { _: [] } });
      assert.fail('Command succeeded while expected fail');
    }
    catch (e: any) {
      assert.strictEqual(e.error.message, 'Error');
    }
  });

  it(`loads all commands, when the matched file doesn't contain command`, async () => {
    sinon.stub(cli, 'loadCommandFromFile').callsFake(_ => (cli.loadCommandFromFile as any).wrappedMethod.apply(cli, [path.join(rootFolder, 'CommandInfo.js')]));
    await cli.loadCommandFromArgs(['status']);

    assert.strictEqual(cli.commandToExecute, undefined);
  });

  it(`loads all commands, when exception was thrown when loading a command file`, async () => {
    sinon.stub(cli, 'loadCommandFromFile').resolves();
    await cli.loadCommandFromArgs(['status']);

    assert.strictEqual(cli.commandToExecute, undefined);
  });

  it('doesn\'t fail when undefined object is passed to the log', async () => {
    const actual = await cli.formatOutput(mockCommand, undefined, { output: 'text' });
    assert.strictEqual(actual, undefined);
  });

  it('returns the same object if non-array is passed to the log', async () => {
    const s = 'foo';
    const actual = await cli.formatOutput(mockCommand, s, { output: 'text' });
    assert.strictEqual(actual, s);
  });

  it('doesn\'t fail when an array with undefined object is passed to the log', async () => {
    const actual = await cli.formatOutput(mockCommand, [undefined], { output: 'text' });
    assert.strictEqual(actual, '');
  });

  it('formats output as pretty JSON when JSON output requested', async () => {
    const o = { lorem: 'ipsum', dolor: 'sit' };
    const actual = await cli.formatOutput(mockCommand, o, { output: 'json' });
    assert.strictEqual(actual, JSON.stringify(o, null, 2));
  });

  it('properly handles new line characters in JSON output', async () => {
    const input = {
      "_ObjectIdentity_": "b61700a0-9062-3000-659e-7f5738e3385a|908bed80-a04a-4433-b4a0-883d9847d110:1b11f502-9eb0-401a-b164-68933e6e9443\nSiteProperties\nhttps%3a%2f%2fm365x954810.sharepoint.com%2fsites%2fsite1617"
    };
    const expected = [
      '{',
      '  "_ObjectIdentity_": "b61700a0-9062-3000-659e-7f5738e3385a|908bed80-a04a-4433-b4a0-883d9847d110:1b11f502-9eb0-401a-b164-68933e6e9443\\\\\\nSiteProperties\\\\\\nhttps%3a%2f%2fm365x954810.sharepoint.com%2fsites%2fsite1617"',
      '}'
    ].join('\n');
    const actual = await cli.formatOutput(mockCommand, input, { output: 'json' });
    assert.strictEqual(actual, expected);
  });

  it('formats object with array as csv', async () => {
    const input =
      [{
        "header1": "value1item1",
        "header2": "value2item1"
      },
      {
        "header1": "value1item2",
        "header2": "value2item2"
      }
      ];
    const expected = "header1,header2\nvalue1item1,value2item1\nvalue1item2,value2item2\n";
    const actual = await cli.formatOutput(mockCommand, input, { output: 'csv' });
    assert.strictEqual(actual, expected);
  });

  it('formats a simple object as csv', async () => {
    const input =
    {
      "header1": "value1item1",
      "header2": "value2item1"
    };
    const expected = "header1,header2\nvalue1item1,value2item1\n";
    const actual = await cli.formatOutput(mockCommand, input, { output: 'csv' });
    assert.strictEqual(actual, expected);
  });

  it('does not produce headers when csvHeader config is set to false ', async () => {
    const input =
    {
      "header1": "value1item1",
      "header2": "value2item1"
    };
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.csvHeader) {
        return false;
      }
      return defaultValue;
    });

    const expected = "value1item1,value2item1\n";
    const actual = await cli.formatOutput(mockCommand, input, { output: 'csv' });
    assert.strictEqual(actual, expected);
  });

  it('quotes all non-empty fields even if not required when csvQuoted config is set to true', async () => {
    const input =
    {
      "header1": "value1item1",
      "header2": "value2item1"
    };
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.csvQuoted) {
        return true;
      }
      return defaultValue;
    });

    const expected = "\"header1\",\"header2\"\n\"value1item1\",\"value2item1\"\n";
    const actual = await cli.formatOutput(mockCommand, input, { output: 'csv' });
    assert.strictEqual(actual, expected);
  });

  it('quotes all empty fields if csvQuotedEmpty config is set to true', async () => {
    const input =
    {
      "header1": "value1item1",
      "header2": ""
    };
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.csvQuotedEmpty) {
        return true;
      }
      return defaultValue;
    });

    const expected = "header1,header2\nvalue1item1,\"\"\n";
    const actual = await cli.formatOutput(mockCommand, input, { output: 'csv' });
    assert.strictEqual(actual, expected);
  });

  it('quotes all fields with character set in csvQuote config', async () => {
    const input =
    {
      "header1": "value1item1",
      "header2": "value2item1"
    };
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.csvQuoted) {
        return true;
      }
      return defaultValue;
    });
    sinon.stub(cli.getConfig(), 'get').callsFake((settingName) => {
      if (settingName === settingsNames.csvQuote) {
        return "_";
      }
      return null;
    });

    const expected = "_header1_,_header2_\n_value1item1_,_value2item1_\n";
    const actual = await cli.formatOutput(mockCommand, input, { output: 'csv' });
    assert.strictEqual(actual, expected);
  });

  it('formats simple output as text', async () => {
    const o = false;
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    assert.strictEqual(actual, `${o}`);
  });

  it('formats date output as text', async () => {
    const d = new Date();
    const actual = await cli.formatOutput(mockCommand, d, { output: 'text' });
    assert.strictEqual(actual, d.toString());
  });

  it('formats object output as transposed table when passing sequential props', async () => {
    const o = { prop1: 'value1', prop2: 'value2' };
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    const t = new Table();
    t.cell('prop1', 'value1');
    t.cell('prop2', 'value2');
    t.newRow();
    const expected = t.printTransposed({
      separator: ': '
    });
    assert.strictEqual(actual, expected);
  });

  it('formats object output as transposed table', async () => {
    const o = { prop1: 'value1 ', prop12: 'value12' };
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    const t = new Table();
    t.cell('prop1', 'value1');
    t.cell('prop12', 'value12');
    t.newRow();
    const expected = t.printTransposed({
      separator: ': '
    });
    assert.strictEqual(actual, expected);
  });

  it('formats array values as JSON', async () => {
    const o = { prop1: ['value1', 'value2'] };
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    const expected = 'prop1: ["value1","value2"]' + '\n';
    assert.strictEqual(actual, expected);
  });

  it('formats array of string arrays output as comma-separated strings', async () => {
    const o = [
      ['value1', 'value2'],
      ['value3', 'value4']
    ];
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    const expected = [o[0].join(','), o[1].join(',')].join(os.EOL);
    assert.strictEqual(actual, expected);
  });

  it('formats array of object output as table', async () => {
    const o = [
      { prop1: 'value1', prop2: 'value2' },
      { prop1: 'value3', prop2: 'value4' }
    ];
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    const t = new Table();
    t.cell('prop1', 'value1');
    t.cell('prop2', 'value2');
    t.newRow();
    t.cell('prop1', 'value3');
    t.cell('prop2', 'value4');
    t.newRow();
    const expected = t.toString();
    assert.strictEqual(actual, expected);
  });

  it('formats command error as error message', async () => {
    const o = new CommandError('An error has occurred');
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    const expected = chalk.red('Error: An error has occurred');
    assert.strictEqual(actual, expected);
  });

  it('sets array type to the first non-undefined value', async () => {
    const o = [undefined, 'lorem', 'ipsum'];
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    const expected = `${os.EOL}lorem${os.EOL}ipsum`;
    assert.strictEqual(actual, expected);
  });

  it('skips primitives mixed with objects when rendering a table', async () => {
    const o = [
      { prop1: 'value1', prop2: 'value2' },
      'lorem',
      { prop1: 'value3', prop2: 'value4' }
    ];
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    const t = new Table();
    t.cell('prop1', 'value1');
    t.cell('prop2', 'value2');
    t.newRow();
    t.cell('prop1', 'value3');
    t.cell('prop2', 'value4');
    t.newRow();
    const expected = t.toString();
    assert.strictEqual(actual, expected);
  });

  it('formats object with array as md', async () => {
    const input =
      [{
        "header1": "value1item1",
        "header2": "value2item1"
      },
      {
        "header1": "value1item2",
        "header2": "value2item2"
      }
      ];
    const actual = await cli.formatOutput(mockCommand, input, { output: 'md' });
    const match = actual.match(/^## /gm);
    assert.strictEqual(match, null);
  });

  it('formats a simple object as md', async () => {
    const input =
    {
      "header1": "value1item1",
      "header2": "value2item1"
    };
    const actual = await cli.formatOutput(mockCommand, input, { output: 'md' });
    const match = actual.match(/^## /gm);
    assert.strictEqual(match, null);
  });

  it('applies JMESPath query to a single object', async () => {
    const o = {
      "first": "Joe",
      "last": "Doe"
    };
    const actual = await cli.formatOutput(mockCommand, o, { query: 'first', output: 'json' });
    assert.strictEqual(actual, JSON.stringify("Joe"));
  });

  it('filters output following command definition in output text', async () => {
    const o = [
      { "name": "Seattle", "state": "WA" },
      { "name": "New York", "state": "NY" },
      { "name": "Bellevue", "state": "WA" },
      { "name": "Olympia", "state": "WA" }
    ];
    cli.commandToExecute = {
      defaultProperties: ['name']
    } as any;
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    const t = new Table();
    t.cell('name', 'Seattle');
    t.newRow();
    t.cell('name', 'New York');
    t.newRow();
    t.cell('name', 'Bellevue');
    t.newRow();
    t.cell('name', 'Olympia');
    t.newRow();
    const expected = t.toString();
    assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected));
  });

  it('filters output wrapped in a value property following command definition in output text', async () => {
    const o = {
      value: [
        { "name": "Seattle", "state": "WA" },
        { "name": "New York", "state": "NY" },
        { "name": "Bellevue", "state": "WA" },
        { "name": "Olympia", "state": "WA" }
      ]
    };
    cli.commandToExecute = {
      defaultProperties: ['name']
    } as any;
    const actual = await cli.formatOutput(mockCommand, o, { output: 'text' });
    const t = new Table();
    t.cell('name', 'Seattle');
    t.newRow();
    t.cell('name', 'New York');
    t.newRow();
    t.cell('name', 'Bellevue');
    t.newRow();
    t.cell('name', 'Olympia');
    t.newRow();
    const expected = t.toString();
    assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected));
  });

  it('applies JMESPath query to an array', async () => {
    const o = {
      "locations": [
        { "name": "Seattle", "state": "WA" },
        { "name": "New York", "state": "NY" },
        { "name": "Bellevue", "state": "WA" },
        { "name": "Olympia", "state": "WA" }
      ]
    };
    const actual = await cli.formatOutput(mockCommand, o, {
      query: `locations[?state == 'WA'].name | sort(@) | {WashingtonCities: join(', ', @)}`,
      output: 'json'
    });
    assert.strictEqual(actual, JSON.stringify({
      "WashingtonCities": "Bellevue, Olympia, Seattle"
    }, null, 2));
  });

  it('doesn\'t apply JMESPath query when command help requested', async () => {
    const o = {
      "locations": [
        { "name": "Seattle", "state": "WA" },
        { "name": "New York", "state": "NY" },
        { "name": "Bellevue", "state": "WA" },
        { "name": "Olympia", "state": "WA" }
      ]
    };
    const actual = await cli.formatOutput(mockCommand, o, {
      query: `locations[?state == 'WA'].name | sort(@) | {WashingtonCities: join(', ', @)}`,
      output: 'json',
      help: true
    });
    assert.strictEqual(actual, JSON.stringify(o, null, 2));
  });

  it('throws human-readable error when invalid JMESPath query specified', async () => {
    const o = {
      "locations": [
        { "name": "Seattle", "state": "WA" },
        { "name": "New York", "state": "NY" },
        { "name": "Bellevue", "state": "WA" },
        { "name": "Olympia", "state": "WA" }
      ]
    };
    assert.rejects(async () => {
      await cli.formatOutput(mockCommand, o, {
        query: `contains(abc)`,
        output: 'json'
      });
    }, chalk.red('Error: JMESPath query error. ArgumentError: contains() takes 2 arguments but received 1. See https://jmespath.org/specification.html for more information'));
  });

  it(`prints commands grouped per service when no command specified`, async () => {
    cli.loadCommandFromArgs(['status']);
    cli.loadCommandFromArgs(['spo', 'site', 'list']);
    cli.printAvailableCommands();

    assert(cliLogStub.calledWith('  cli *  8 commands'));
  });

  it(`prints commands from the specified group`, async () => {
    cli.loadCommandFromArgs(['status']);
    cli.loadCommandFromArgs(['spo', 'site', 'list']);
    cli.optionsFromArgs = {
      options: {
        _: ['cli']
      }
    };
    cli.printAvailableCommands();

    assert(cliLogStub.calledWith('  cli mock *        5 commands'));
  });

  it(`prints commands from the root group when the specified string doesn't match any group`, () => {
    cli.loadCommandFromArgs(['status']);
    cli.loadCommandFromArgs(['spo', 'site', 'list']);
    cli.optionsFromArgs = {
      options: {
        _: ['foo']
      }
    };
    cli.printAvailableCommands();

    assert(cliLogStub.calledWith('  cli *  8 commands'));
  });

  it(`runs properly when context file not found`, async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '--parameterX', '123', '--output', 'json']);
    assert(cliLogStub.called);
  });

  it(`populates option from context file`, async () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString() === '.m365rc.json');
    sinon.stub(fs, 'readFileSync').onCall(0).returns('{"context": {"parameterY": "456"}}').onCall(1).returns('{}');
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return undefined;
      }
      return defaultValue;
    });
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '--parameterX', '123', '--output', 'text']);
    assert(cliLogStub.called);
  });

  it(`populates option from context file (debug mode)`, async () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString() === '.m365rc.json');
    sinon.stub(fs, 'readFileSync').onCall(0).returns('{"context": {"parameterY": "456"}}').onCall(1).returns('{}');
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return undefined;
      }
      return defaultValue;
    });
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '--parameterX', '123', '--output', 'text', '--debug']);
    assert(cliLogStub.called);
  });

  it(`runs properly when context m365rc file found but without any context`, async () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString() === '.m365rc.json');
    sinon.stub(fs, 'readFileSync').onCall(0).returns('{}').onCall(1).returns('{}');
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return undefined;
      }
      return defaultValue;
    });
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '--parameterX', '123', '--output', 'text']);
    assert(cliLogStub.called);
  });

  it(`throws error when context json parse fails`, async () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString() === '.m365rc.json');
    sinon.stub(fs, 'readFileSync').onCall(0).returns('I will not parse').onCall(1).returns('{}');
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return undefined;
      }
      return defaultValue;
    });
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');
    try {
      await cli.execute(['cli', 'mock', '--parameterX', '123', '--output', 'text']);
      assert.fail(`Didn't fail while expected`);
    }
    catch {
      assert(cliErrorStub.calledWith(chalk.red('Error: Error parsing .m365rc.json')));
    }
  });

  it(`exits with the specified exit code`, async () => {
    try {
      await cli.closeWithError(new CommandError('Error', 5), { options: {} });
      assert.fail(`Didn't fail while expected`);
    }
    catch {
      assert(processExitStub.calledWith(5));
    }
  });

  it(`prints error as JSON in JSON output mode and printErrorsAsPlainText set to false`, async () => {
    const config = cli.getConfig();
    sinon.stub(config, 'get').callsFake(() => false);

    try {
      await cli.closeWithError(new CommandError('Error'), { options: { output: 'json' } });
      assert.fail(`Didn't fail while expected`);
    }
    catch (err) {
      assert(cliErrorStub.calledWith(JSON.stringify({ error: 'Error' })));
    }
  });

  it(`replaces option value with the content of the specified file when value starts with @ and the specified file exists`, async () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().endsWith('.txt'));
    sinon.stub(fs, 'readFileSync').returns('abc');
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '-x', '@file.txt', '-o', 'text']);
    assert(cliLogStub.calledWith('abc'));
  });

  it(`returns error when reading file contents failed`, async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').throws(new CommandError('An error has occurred'));
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    try {
      await cli.execute(['cli', 'mock', '-x', '@file.txt']);
      assert.fail('Promise completed while error expected');
    }
    catch {
      assert(cliErrorStub.calledWith(chalk.red(`Error: An error has occurred`)));
    }
  });

  it(`leaves the original value if the file specified in @ value doesn't exist`, async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    await cli.execute(['cli', 'mock', '-x', '@file.txt', '-o', 'text']);
    assert(cliLogStub.calledWith('@file.txt'));
  });

  it(`closes with error when processing options failed`, async () => {
    sinon.stub(mockCommand, 'processOptions').rejects('Error');
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock');

    try {
      await cli.execute(['cli', 'mock', '-x', '123']);
      assert.fail('Promise completed while error expected');
    }
    catch (e: any) {
      assert.strictEqual(e.name, 'Error');
    }
  });

  it(`logs output to console`, () => {
    sinonUtil.restore(cli.log);
    const consoleLogSpy: sinon.SinonSpy = sinon.stub(console, 'log').returns();
    cli.log('Message');
    assert(consoleLogSpy.calledWith('Message'));
  });

  it(`logs empty line to console when no message specified`, () => {
    sinonUtil.restore(cli.log);
    const consoleLogSpy: sinon.SinonSpy = sinon.stub(console, 'log').returns();
    cli.log();
    assert(consoleLogSpy.calledWith());
  });

  it(`logs error to console stderr`, async () => {
    sinonUtil.restore(cli.error);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((_, defaultValue) => defaultValue);
    const consoleErrorStub = sinon.stub(console, 'error').returns();

    await cli.error('Message');
    assert(consoleErrorStub.calledWith('Message'));
  });

  it(`logs error to console stdout when stdout configured as error output`, async () => {
    const config = cli.getConfig();
    sinon.stub(config, 'get').returns('stdout');
    sinonUtil.restore(cli.error);
    const consoleErrorSpy: sinon.SinonSpy = sinon.stub(console, 'error').returns();
    const consoleLogSpy: sinon.SinonSpy = sinon.stub(console, 'log').returns();

    await cli.error('Message');
    assert(consoleErrorSpy.notCalled, 'console.error called');
    assert(consoleLogSpy.calledWith('Message'), 'console.log not called with the right message');
  });

  it(`returns stored configuration value when available`, () => {
    const config = cli.getConfig();
    sinon.stub(config, 'get').callsFake(() => 'value');
    const actualValue = cli.getSettingWithDefaultValue('key', '');
    assert.strictEqual(actualValue, 'value');
  });

  it('returns true, for the method shouldTrimOutput, when output is text', () => {
    const spyShouldTrimOutput = cli.shouldTrimOutput('text');
    assert.strictEqual(spyShouldTrimOutput, true);
  });

  it('returns false, for the method shouldTrimOutput, when output is csv', () => {
    const spyShouldTrimOutput = cli.shouldTrimOutput('csv');
    assert.strictEqual(spyShouldTrimOutput, false);
  });

  it('returns false, for the method shouldTrimOutput, when output is json', () => {
    const spyShouldTrimOutput = cli.shouldTrimOutput('json');
    assert.strictEqual(spyShouldTrimOutput, false);
  });

  it('does not show help when output is set to none', async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli completion clink update');

    await cli.execute(['cli', 'completion', 'clink', 'update', '-h', 'examples', '--output', 'none']);
    assert.strictEqual(log.length === 0, true);
  });

  it(`shows no output on successful run with output set to none`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock boolean rewrite');

    await cli.execute(['cli', 'mock', 'boolean', 'rewrite', '--booleanParameterX', 'true', '--booleanParameterY', 'false', '--output', 'none']);
    assert.strictEqual(cliLogStub.notCalled, true);
    assert.strictEqual(cliErrorStub.notCalled, true);
  });

  it(`shows no output when a validation error occurs in and output is set to none`, async () => {
    cli.commandToExecute = cli.commands.find(c => c.name === 'cli mock boolean rewrite');

    await cli.execute(['cli', 'mock', 'boolean', 'rewrite', '--booleanParameterX', 'folse', '--output', 'none']);
    assert(cliErrorStub.notCalled);
    assert(cliLogStub.notCalled);
  });

  it('for completion commands loads full command info', async () => {
    sinonUtil.restore(cli.loadAllCommandsInfo);
    const loadAllCommandsInfoStub = sinon.spy(cli, 'loadAllCommandsInfo');
    sinon.stub(cli, 'executeCommand').callsFake(() => Promise.resolve());

    await cli.execute(['cli', 'completion', 'sh', 'update']);
    assert(loadAllCommandsInfoStub.calledWith(true));
  });
});