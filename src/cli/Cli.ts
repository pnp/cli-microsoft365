import Configstore from 'configstore';
import fs from 'fs';
import minimist from 'minimist';
import { createRequire } from 'module';
import ora, { Options, Ora } from 'ora';
import os from 'os';
import path from 'path';
import { fileURLToPath, pathToFileURL } from 'url';
import Command, { CommandArgs, CommandError, CommandTypes } from '../Command.js';
import GlobalOptions from '../GlobalOptions.js';
import config from '../config.js';
import { M365RcJson } from '../m365/base/M365RcJson.js';
import request from '../request.js';
import { settingsNames } from '../settingsNames.js';
import { telemetry } from '../telemetry.js';
import { app } from '../utils/app.js';
import { formatting } from '../utils/formatting.js';
import { md } from '../utils/md.js';
import { validation } from '../utils/validation.js';
import { CommandInfo } from './CommandInfo.js';
import { CommandOptionInfo } from './CommandOptionInfo.js';
import { Logger } from './Logger.js';
import { SelectionConfig, ConfirmationConfig, prompt } from '../utils/prompt.js';
import { timings } from './timings.js';
const require = createRequire(import.meta.url);

const __dirname = fileURLToPath(new URL('.', import.meta.url));

export interface CommandOutput {
  stdout: string;
  stderr: string;
}

export class Cli {
  public commands: CommandInfo[] = [];
  /**
   * Command to execute
   */
  public commandToExecute: CommandInfo | undefined;
  /**
   * Name of the command specified through args
   */
  public currentCommandName: string | undefined;
  private optionsFromArgs: { options: minimist.ParsedArgs } | undefined;
  private static instance: Cli;
  private static defaultHelpMode = 'options';
  public static helpModes: string[] = ['options', 'examples', 'remarks', 'response', 'full'];
  public spinner: Ora;

  private _config: Configstore | undefined;
  public get config(): Configstore {
    if (!this._config) {
      this._config = new Configstore(config.configstoreName);
    }

    return this._config;
  }

  public getSettingWithDefaultValue<TValue>(settingName: string, defaultValue: TValue): TValue {
    const configuredValue: TValue | undefined = this.config.get(settingName);
    if (typeof configuredValue === 'undefined') {
      return defaultValue;
    }
    else {
      return configuredValue;
    }
  }

  private constructor() {
    const options: Options = {
      text: 'Running command...',
      /* c8 ignore next 1 */
      stream: this.getSettingWithDefaultValue('errorOutput', 'stderr') === 'stderr' ? process.stderr : process.stdout
    };
    this.spinner = ora(options);
  }

  public static getInstance(): Cli {
    if (!Cli.instance) {
      Cli.instance = new Cli();
    }

    return Cli.instance;
  }

  public async execute(rawArgs: string[]): Promise<void> {
    const start = process.hrtime.bigint();

    // for completion commands we also need information about commands' options
    const loadAllCommandInfo: boolean = rawArgs.indexOf('completion') > -1;
    this.loadAllCommandsInfo(loadAllCommandInfo);

    // check if help for a specific command has been requested using the
    // 'm365 help xyz' format. If so, remove 'help' from the array of words
    // to use lazy loading commands but keep track of the fact that help should
    // be displayed
    let showHelp: boolean = false;
    if (rawArgs.length > 0 && rawArgs[0] === 'help') {
      showHelp = true;
      rawArgs.shift();
    }

    // parse args to see if a command has been specified
    const parsedArgs: minimist.ParsedArgs = minimist(rawArgs);

    // load command
    await this.loadCommandFromArgs(parsedArgs._);

    if (this.commandToExecute) {
      // we have found a command to execute. Parse args again taking into
      // account short and long options, option types and whether the command
      // supports known and unknown options or not

      try {
        this.optionsFromArgs = {
          options: this.getCommandOptionsFromArgs(rawArgs, this.commandToExecute)
        };
      }
      catch (e: any) {
        const optionsWithoutShorts = Cli.removeShortOptions({ options: parsedArgs });

        return this.closeWithError(e.message, optionsWithoutShorts, false);
      }
    }
    else {
      // we need this to properly support displaying commands
      // from the current group
      this.optionsFromArgs = {
        options: parsedArgs
      };
    }

    // show help if no match found, help explicitly requested or
    // no command specified
    if (!this.commandToExecute ||
      showHelp ||
      parsedArgs.h ||
      parsedArgs.help) {
      if (parsedArgs.output !== 'none') {
        this.printHelp(await this.getHelpMode(parsedArgs));
      }
      return Promise.resolve();
    }

    delete (this.optionsFromArgs.options as any)._;
    delete (this.optionsFromArgs.options as any)['--'];

    try {
      // replace values staring with @ with file contents
      Cli.loadOptionValuesFromFiles(this.optionsFromArgs);
    }
    catch (e) {
      return this.closeWithError(e, this.optionsFromArgs);
    }

    const startProcessing = process.hrtime.bigint();
    try {
      // process options before passing them on to validation stage
      const contextCommandOptions = await this.loadOptionsFromContext(this.commandToExecute.options, this.optionsFromArgs.options.debug);
      this.optionsFromArgs.options = { ...contextCommandOptions, ...this.optionsFromArgs.options };
      await this.commandToExecute.command.processOptions(this.optionsFromArgs.options);

      const endProcessing = process.hrtime.bigint();
      timings.options.push(Number(endProcessing - startProcessing));
    }
    catch (e: any) {
      const endProcessing = process.hrtime.bigint();
      timings.options.push(Number(endProcessing - startProcessing));

      return this.closeWithError(e.message, this.optionsFromArgs, false);
    }

    // if output not specified, set the configured output value (if any)
    if (this.optionsFromArgs.options.output === undefined) {
      this.optionsFromArgs.options.output = this.getSettingWithDefaultValue<string | undefined>(settingsNames.output, 'json');
    }

    const startValidation = process.hrtime.bigint();
    const validationResult = await this.commandToExecute.command.validate(this.optionsFromArgs, this.commandToExecute);
    const endValidation = process.hrtime.bigint();
    timings.validation.push(Number(endValidation - startValidation));
    if (validationResult !== true) {
      return this.closeWithError(validationResult, this.optionsFromArgs, true);
    }

    const end = process.hrtime.bigint();
    timings.core.push(Number(end - start));

    try {
      await Cli.executeCommand(this.commandToExecute.command, this.optionsFromArgs);
      const endTotal = process.hrtime.bigint();
      timings.total.push(Number(endTotal - start));
      this.printTimings(rawArgs);
      process.exit(0);
    }
    catch (err) {
      const endTotal = process.hrtime.bigint();
      timings.total.push(Number(endTotal - start));
      this.printTimings(rawArgs);
      await this.closeWithError(err, this.optionsFromArgs);
      /* c8 ignore next */
    }
  }

  private printTimings(rawArgs: string[]): void {
    if (rawArgs.some(arg => arg === '--debug')) {
      Cli.error('');
      Cli.error('Timings:');
      Object.getOwnPropertyNames(timings).forEach(key => {
        Cli.error(`${key}: ${(timings as any)[key].reduce((a: number, b: number) => a + b, 0) / 1e6}ms`);
      });
    }
  }

  public static async executeCommand(command: Command, args: { options: minimist.ParsedArgs }): Promise<void> {
    const logger: Logger = {
      log: async (message: any): Promise<void> => {
        if (args.options.output !== 'none') {
          const output: any = await Cli.formatOutput(command, message, args.options);
          Cli.log(output);
        }
      },
      logRaw: async (message: any): Promise<void> => {
        if (args.options.output !== 'none') {
          Cli.log(message);
        }
      },
      logToStderr: async (message: any): Promise<void> => {
        if (args.options.output !== 'none') {
          Cli.error(message);
        }
      }
    };

    if (args.options.debug) {
      await logger.logToStderr(`Executing command ${command.name} with options ${JSON.stringify(args)}`);
    }

    // store the current command name, if any and set the name to the name of
    // the command to execute
    const cli = Cli.getInstance();
    const parentCommandName: string | undefined = cli.currentCommandName;
    cli.currentCommandName = command.getCommandName(cli.currentCommandName);
    const showSpinner = cli.getSettingWithDefaultValue<boolean>(settingsNames.showSpinner, true) && args.options.output !== 'none';

    // don't show spinner if running tests
    /* c8 ignore next 3 */
    if (showSpinner && typeof global.it === 'undefined') {
      cli.spinner.start();
    }

    const startCommand = process.hrtime.bigint();
    try {
      await command.action(logger, args as any);

      if (args.options.debug || args.options.verbose) {
        const chalk = (await import('chalk')).default;
        await logger.logToStderr(chalk.green('DONE'));
      }
    }
    finally {
      // restore the original command name
      cli.currentCommandName = parentCommandName;

      /* c8 ignore next 3 */
      if (cli.spinner.isSpinning) {
        cli.spinner.stop();
      }

      const endCommand = process.hrtime.bigint();
      timings.command.push(Number(endCommand - startCommand));
    }
  }

  public static async executeCommandWithOutput(command: Command, args: { options: minimist.ParsedArgs }, listener?: {
    stdout?: (message: any) => void,
    stderr?: (message: any) => void
  }): Promise<CommandOutput> {
    const log: string[] = [];
    const logErr: string[] = [];
    const logger: Logger = {
      log: async (message: any): Promise<void> => {
        const formattedMessage = await Cli.formatOutput(command, message, args.options);
        if (listener && listener.stdout) {
          listener.stdout(formattedMessage);
        }
        log.push(formattedMessage);
      },
      logRaw: async (message: any): Promise<void> => {
        const formattedMessage = await Cli.formatOutput(command, message, args.options);
        if (listener && listener.stdout) {
          listener.stdout(formattedMessage);
        }
        log.push(formattedMessage);
      },
      logToStderr: async (message: any): Promise<void> => {
        if (listener && listener.stderr) {
          listener.stderr(message);
        }
        logErr.push(message);
      }
    };

    if (args.options.debug && args.options.output !== 'none') {
      const message = `Executing command ${command.name} with options ${JSON.stringify(args)}`;
      if (listener && listener.stderr) {
        listener.stderr(message);
      }
      logErr.push(message);
    }

    // store the current command name, if any and set the name to the name of
    // the command to execute
    const cli = Cli.getInstance();
    const parentCommandName: string | undefined = cli.currentCommandName;
    cli.currentCommandName = command.getCommandName();
    // store the current logger if any
    const currentLogger: Logger | undefined = request.logger;

    try {
      await command.action(logger, args as any);

      return ({
        stdout: log.join(os.EOL),
        stderr: logErr.join(os.EOL)
      });
    }
    catch (err: any) {
      // restoring the command and logger is done here instead of in a 'finally' because there were issues with the code coverage tool
      // restore the original command name
      cli.currentCommandName = parentCommandName;
      // restore the original logger
      request.logger = currentLogger;

      throw {
        error: err,
        stderr: logErr.join(os.EOL)
      };
    }
    /* c8 ignore next */
    finally {
      // restore the original command name
      cli.currentCommandName = parentCommandName;
      // restore the original logger
      request.logger = currentLogger;
    }
  }

  public loadAllCommandsInfo(loadFull: boolean = false): void {
    const commandsInfoFileName = loadFull ? 'allCommandsFull.json' : 'allCommands.json';
    this.commands = require(path.join(__dirname, '..', '..', commandsInfoFileName));
  }

  /**
   * Loads command files into CLI based on the specified arguments.
   *
   * @param commandNameWords Array of words specified as args
  */
  public async loadCommandFromArgs(commandNameWords: string[]): Promise<void> {
    if (commandNameWords.length === 0) {
      return;
    }

    this.currentCommandName = commandNameWords.join(' ');

    const commandFilePath = this.commands
      .find(c => c.name === this.currentCommandName ||
        c.aliases?.find(a => a === this.currentCommandName))?.file ?? '';

    if (commandFilePath) {
      await this.loadCommandFromFile(commandFilePath);
    }
  }

  private async loadOptionsFromContext(commandOptions: CommandOptionInfo[], debug: boolean | undefined): Promise<any> {
    const filePath: string = '.m365rc.json';
    let m365rc: M365RcJson = {};

    if (!fs.existsSync(filePath)) {
      return;
    }

    if (debug!) {
      await Cli.error('found .m365rc.json file');
    }

    try {
      const fileContents: string = fs.readFileSync(filePath, 'utf8');
      if (fileContents) {
        m365rc = JSON.parse(fileContents);
      }
    }
    catch (e) {
      await this.closeWithError(`Error parsing ${filePath}`, { options: {} });
      /* c8 ignore next */
    }

    if (!m365rc.context) {
      return;
    }

    if (debug!) {
      await Cli.error('found context in .m365rc.json file');
    }

    const context = m365rc.context;

    const foundOptions: any = {};
    await commandOptions.forEach(async option => {
      if (context[option.name]) {
        foundOptions[option.name] = context[option.name];
        if (debug!) {
          await Cli.error(`returning ${option.name} option from context`);
        }
      }
    });

    return foundOptions;
  }

  /**
   * Loads command from the specified file into CLI. If can't find the file
   * or the file doesn't contain a command, loads all available commands.
   *
   * @param commandFilePathUrl File path of the file with command to load
   */
  private async loadCommandFromFile(commandFileUrl: string): Promise<void> {
    const commandsFolder = path.join(__dirname, '../m365');
    const filePath: string = path.join(commandsFolder, commandFileUrl);

    if (!fs.existsSync(filePath)) {
      // reset command name
      this.currentCommandName = undefined;
      return;
    }

    try {
      const command: any = await import(pathToFileURL(filePath).toString());
      if (command.default instanceof Command) {
        const commandInfo = this.commands.find(c => c.file === commandFileUrl);
        this.commandToExecute = Cli.getCommandInfo(command.default, commandFileUrl, commandInfo?.help);
      }
    }
    catch { }
  }

  public static getCommandInfo(command: Command, filePath: string = '', helpFilePath: string = ''): CommandInfo {
    return {
      aliases: command.alias(),
      name: command.name,
      description: command.description,
      command: command,
      options: this.getCommandOptions(command),
      defaultProperties: command.defaultProperties(),
      file: filePath,
      help: helpFilePath
    };
  }

  private static getCommandOptions(command: Command): CommandOptionInfo[] {
    const options: CommandOptionInfo[] = [];

    command.options.forEach(option => {
      const required: boolean = option.option.indexOf('<') > -1;
      const optionArgs: string[] = option.option.split(/[ ,|]+/);
      let short: string | undefined;
      let long: string | undefined;
      let name: string = '';
      optionArgs.forEach(o => {
        if (o.startsWith('--')) {
          long = o.replace('--', '');
          name = long;
        }
        else if (o.startsWith('-')) {
          short = o.replace('-', '');
          name = short;
        }
      });

      options.push({
        autocomplete: option.autocomplete,
        long: long,
        name: name,
        required: required,
        short: short
      });
    });

    return options;
  }

  private getCommandOptionsFromArgs(args: string[], commandInfo: CommandInfo | undefined): minimist.ParsedArgs {
    const minimistOptions: minimist.Opts = {
      alias: {}
    };

    let argsToParse = args;

    if (commandInfo) {
      const commandTypes = commandInfo.command.types;
      if (commandTypes) {
        minimistOptions.string = commandTypes.string;

        // minimist will parse unused boolean options to 'false' (unused options => options that are not included in the args)
        // But in the CLI booleans are nullable. They can can be true, false or undefined.
        // For this reason we only pass boolean types that are actually used as arg.
        minimistOptions.boolean = commandTypes.boolean.filter(optionName => args.some(arg => `--${optionName}` === arg || `-${optionName}` === arg));
      }

      minimistOptions.alias = {};
      commandInfo.options.forEach(option => {
        if (option.short && option.long) {
          (minimistOptions.alias as any)[option.short] = option.long;
        }
      });

      argsToParse = this.getRewrittenArgs(args, commandTypes);
    }

    return minimist(argsToParse, minimistOptions);
  }

  /**
   * Rewrites arguments (if necessary) before passing them into minimist.
   * Currently only boolean values are checked and fixed.
   * Args are only checked and rewritten if the option has been added to the 'types.boolean' array.
   */
  private getRewrittenArgs(args: string[], commandTypes: CommandTypes): string[] {
    const booleanTypes = commandTypes.boolean;

    if (booleanTypes.length === 0) {
      return args;
    }

    return args.map((arg: string, index: number, array: string[]) => {
      if (arg.startsWith('-') || index === 0) {
        return arg;
      }

      // This line checks if the current arg is a value that belongs to a boolean option.
      if (booleanTypes.some(t => `--${t}` === array[index - 1] || `-${t}` === array[index - 1])) {
        const rewrittenBoolean = formatting.rewriteBooleanValue(arg);

        if (!validation.isValidBoolean(rewrittenBoolean)) {
          const optionName = array[index - 1];
          throw new Error(`The value '${arg}' for option '${optionName}' is not a valid boolean`);
        }

        return rewrittenBoolean;
      }

      return arg;
    });
  }

  private static async formatOutput(command: Command, logStatement: any, options: GlobalOptions): Promise<any> {
    if (logStatement instanceof Date) {
      return logStatement.toString();
    }

    let logStatementType: string = typeof logStatement;

    if (logStatementType === 'undefined') {
      return logStatement;
    }

    // we need to get the list of object's properties to see if the specified
    // JMESPath query (if any) filters object's properties or not. We need to
    // know this in order to decide if we should use default command's
    // properties or custom ones from JMESPath
    const originalObject: any = Array.isArray(logStatement) ? Cli.getFirstNonUndefinedArrayItem(logStatement) : logStatement;
    const originalProperties: string[] = originalObject && typeof originalObject !== 'string' ? Object.getOwnPropertyNames(originalObject) : [];

    if (options.query &&
      !options.help) {
      const jmespath = await import('jmespath');
      try {
        logStatement = jmespath.search(logStatement, options.query);
      }
      catch (e: any) {
        const message = `JMESPath query error. ${e.message}. See https://jmespath.org/specification.html for more information`;
        await Cli.getInstance().closeWithError(message, { options }, false);
        /* c8 ignore next */
      }
      // we need to update the statement type in case the JMESPath query
      // returns an object of different shape than the original message to log
      // #2095
      logStatementType = typeof logStatement;
    }

    if (!options.output || options.output === 'json') {
      return command.getJsonOutput(logStatement);
    }

    if (logStatement instanceof CommandError) {
      const chalk = (await import('chalk')).default;
      return chalk.red(`Error: ${logStatement.message}`);
    }

    let arrayType: string = '';
    if (!Array.isArray(logStatement)) {
      logStatement = [logStatement];
      arrayType = logStatementType;
    }
    else {
      for (let i: number = 0; i < logStatement.length; i++) {
        if (Array.isArray(logStatement[i])) {
          arrayType = 'array';
          break;
        }

        const t: string = typeof logStatement[i];
        if (t !== 'undefined') {
          arrayType = t;
          break;
        }
      }
    }

    if (arrayType !== 'object') {
      return logStatement.join(os.EOL);
    }

    // if output type has been set to 'text' or 'csv', process the retrieved
    // data so that returned objects contain only default properties specified
    // on the current command. If there is no current command or the
    // command doesn't specify default properties, return original data
    if (this.shouldTrimOutput(options.output)) {
      const cli: Cli = Cli.getInstance();
      const currentCommand: CommandInfo | undefined = cli.commandToExecute;

      if (arrayType === 'object' &&
        currentCommand && currentCommand.defaultProperties) {
        // the log statement contains the same properties as the original object
        // so it can be filtered following the default properties specified on
        // the command
        if (JSON.stringify(originalProperties) === JSON.stringify(Object.getOwnPropertyNames(logStatement[0]))) {
          // in some cases we return properties wrapped in `value` array
          // returned by the API. We'll remove it in the future, but for now
          // we'll use a workaround to drop the `value` array here
          if (logStatement[0].value &&
            Array.isArray(logStatement[0].value)) {
            logStatement = logStatement[0].value;
          }

          logStatement = logStatement.map((s: any) =>
            formatting.filterObject(s, currentCommand.defaultProperties as string[]));
        }
      }
    }

    switch (options.output) {
      case 'csv':
        return command.getCsvOutput(logStatement, options);
      case 'md':
        return command.getMdOutput(logStatement, command, options);
      default:
        return command.getTextOutput(logStatement);
    }
  }

  private static getFirstNonUndefinedArrayItem(arr: any[]): any {
    for (let i: number = 0; i < arr.length; i++) {
      const a: any = arr[i];
      if (typeof a !== 'undefined') {
        return a;
      }
    }

    return undefined;
  }

  private printHelp(helpMode: string, exitCode: number = 0): void {
    const properties: any = {};

    if (this.commandToExecute) {
      properties.command = this.commandToExecute.name;
      this.printCommandHelp(helpMode);
    }
    else {
      Cli.log();
      Cli.log(`CLI for Microsoft 365 v${app.packageJson().version}`);
      Cli.log(`${app.packageJson().description}`);
      Cli.log();

      properties.command = 'commandList';
      this.printAvailableCommands();
    }

    telemetry.trackEvent('help', properties);

    process.exit(exitCode);
  }

  private printCommandHelp(helpMode: string): void {
    const docsRootDir = path.join(__dirname, '..', '..', 'docs');
    const helpFilePath = path.join(docsRootDir, 'docs', 'cmd', this.commandToExecute!.help!);

    if (fs.existsSync(helpFilePath)) {
      let helpContents = fs.readFileSync(helpFilePath, 'utf8');
      helpContents = this.getHelpSection(helpMode, helpContents);
      helpContents = md.md2plain(helpContents, docsRootDir);
      Cli.log();
      Cli.log(helpContents);
    }
  }

  private async getHelpMode(options: any): Promise<string> {
    const { h, help } = options;

    if (!h && !help) {
      return this.getSettingWithDefaultValue<string>(settingsNames.helpMode, Cli.defaultHelpMode);
    }

    // user passed -h or --help, let's see if they passed a specific mode
    // or requested the default
    const helpMode: boolean | string = h ?? help;

    if (typeof helpMode === 'boolean' || typeof helpMode !== 'string') {
      // requested default mode or passed a number, let's use default
      return this.getSettingWithDefaultValue<string>(settingsNames.helpMode, Cli.defaultHelpMode);
    }
    else {
      const lowerCaseHelpMode = helpMode.toLowerCase();

      if (Cli.helpModes.indexOf(lowerCaseHelpMode) < 0) {
        await Cli.getInstance().closeWithError(`Unknown help mode ${helpMode}. Allowed values are ${Cli.helpModes.join(', ')}`, { options }, false);
        /* c8 ignore next 2 */
        return ''; // noop
      }
      else {
        return lowerCaseHelpMode;
      }
    }
  }

  private getHelpSection(helpMode: string, helpContents: string): string {
    if (helpMode === 'full') {
      return helpContents;
    }

    // options is the first section, so get help up to options
    const titleAndUsage = helpContents.substring(0, helpContents.indexOf('## Options'));

    // find the requested section
    const sectionLines: string[] = [];
    const sectionName = helpMode[0].toUpperCase() + helpMode.substring(1);
    const lines: string[] = helpContents.split('\n');
    for (let i: number = 0; i < lines.length; i++) {
      const line = lines[i];

      if (line.indexOf(`## ${sectionName}`) === 0) {
        sectionLines.push(line);
      }
      else if (sectionLines.length > 0) {
        if (line.indexOf('## ') === 0) {
          // we've reached the next section, stop
          break;
        }
        else {
          sectionLines.push(line);
        }
      }
    }

    return titleAndUsage + sectionLines.join('\n');
  }

  private printAvailableCommands(): void {
    // commands that match the current group
    const commandsToPrint: { [commandName: string]: CommandInfo } = {};
    // sub-commands in the current group
    const commandGroupsToPrint: { [group: string]: number } = {};
    // current command group, eg. 'spo', 'spo site'
    let currentGroup: string = '';

    const addToList = (commandName: string, command: CommandInfo): void => {
      const pos: number = commandName.indexOf(' ', currentGroup.length + 1);
      if (pos === -1) {
        commandsToPrint[commandName] = command;
      }
      else {
        const subCommandsGroup: string = commandName.substr(0, pos);
        if (!commandGroupsToPrint[subCommandsGroup]) {
          commandGroupsToPrint[subCommandsGroup] = 0;
        }

        commandGroupsToPrint[subCommandsGroup]++;
      }
    };

    // get current command group
    if (this.optionsFromArgs &&
      this.optionsFromArgs.options &&
      this.optionsFromArgs.options._ &&
      this.optionsFromArgs.options._.length > 0) {
      currentGroup = this.optionsFromArgs.options._.join(' ');

      if (currentGroup) {
        currentGroup += ' ';
      }
    }

    const getCommandsForGroup = (): void => {
      for (let i = 0; i < this.commands.length; i++) {
        const command: CommandInfo = this.commands[i];
        if (command.name.startsWith(currentGroup)) {
          addToList(command.name, command);
        }

        if (command.aliases) {
          for (let j = 0; j < command.aliases.length; j++) {
            const alias: string = command.aliases[j];
            if (alias.startsWith(currentGroup)) {
              addToList(alias, command);
            }
          }
        }
      }
    };

    getCommandsForGroup();
    if (Object.keys(commandsToPrint).length === 0 &&
      Object.keys(commandGroupsToPrint).length === 0) {
      // specified string didn't match any commands. Reset group and try again
      currentGroup = '';
      getCommandsForGroup();
    }

    const namesOfCommandsToPrint: string[] = Object.keys(commandsToPrint);
    if (namesOfCommandsToPrint.length > 0) {
      // determine the length of the longest command name to pad strings + ' [options]'
      const maxLength: number = Math.max(...namesOfCommandsToPrint.map(s => s.length)) + 10;

      Cli.log(`Commands:`);
      Cli.log();

      const sortedCommandNamesToPrint = Object.getOwnPropertyNames(commandsToPrint).sort();
      sortedCommandNamesToPrint.forEach(commandName => {
        Cli.log(`  ${`${commandName} [options]`.padEnd(maxLength, ' ')}  ${commandsToPrint[commandName].description}`);
      });
    }

    const namesOfCommandGroupsToPrint: string[] = Object.keys(commandGroupsToPrint);
    if (namesOfCommandGroupsToPrint.length > 0) {
      if (namesOfCommandsToPrint.length > 0) {
        Cli.log();
      }

      // determine the longest command group name to pad strings + ' *'
      const maxLength: number = Math.max(...namesOfCommandGroupsToPrint.map(s => s.length)) + 2;

      Cli.log(`Commands groups:`);
      Cli.log();

      // sort commands groups (because of aliased commands)
      const sortedCommandGroupsToPrint = Object
        .keys(commandGroupsToPrint)
        .sort()
        .reduce((object: { [group: string]: number }, key: string) => {
          object[key] = commandGroupsToPrint[key];
          return object;
        }, {});

      for (const commandGroup in sortedCommandGroupsToPrint) {
        Cli.log(`  ${`${commandGroup} *`.padEnd(maxLength, ' ')}  ${commandGroupsToPrint[commandGroup]} command${commandGroupsToPrint[commandGroup] === 1 ? '' : 's'}`);
      }
    }

    Cli.log();
  }

  private async closeWithError(error: any, args: CommandArgs, showHelpIfEnabled: boolean = false): Promise<void> {
    let exitCode: number = 1;

    if (args.options.output === 'none') {
      return process.exit(exitCode);
    }

    const chalk = (await import('chalk')).default;

    let errorMessage: string = error instanceof CommandError ? error.message : error;
    if ((!args.options.output || args.options.output === 'json') &&
      !this.getSettingWithDefaultValue<boolean>(settingsNames.printErrorsAsPlainText, true)) {
      errorMessage = JSON.stringify({ error: errorMessage });
    }
    else {
      errorMessage = chalk.red(`Error: ${errorMessage}`);
    }

    if (error instanceof CommandError && error.code) {
      exitCode = error.code;
    }

    await Cli.error(errorMessage);

    if (showHelpIfEnabled &&
      await this.getSettingWithDefaultValue<boolean>(settingsNames.showHelpOnFailure, showHelpIfEnabled)) {
      this.printHelp(await this.getHelpMode(args.options), exitCode);
    }
    else {
      process.exit(exitCode);
    }

    // will never be run. Required for testing where we're stubbing process.exit
    /* c8 ignore next */
    throw new Error(errorMessage);
    /* c8 ignore next */
  }

  public static log(message?: any, ...optionalParams: any[]): void {
    const cli = Cli.getInstance();
    const spinnerSpinning = cli.spinner.isSpinning;

    /* c8 ignore next 3 */
    if (spinnerSpinning) {
      cli.spinner.stop();
    }

    if (message) {
      console.log(message, ...optionalParams);
    }
    else {
      console.log();
    }

    // Restart the spinner if it was running before the log
    /* c8 ignore next 3 */
    if (spinnerSpinning) {
      cli.spinner.start();
    }
  }

  public static async error(message?: any, ...optionalParams: any[]): Promise<void> {
    const cli = Cli.getInstance();
    const spinnerSpinning = cli.spinner.isSpinning;

    /* c8 ignore next 3 */
    if (spinnerSpinning) {
      cli.spinner.stop();
    }

    const errorOutput: string = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');
    if (errorOutput === 'stdout') {
      console.log(message, ...optionalParams);
    }
    else {
      console.error(message, ...optionalParams);
    }

    // Restart the spinner if it was running before the log
    /* c8 ignore next 3 */
    if (spinnerSpinning) {
      cli.spinner.start();
    }
  }

  public static async promptForSelection<T>(config: SelectionConfig<T>): Promise<T> {
    const cli = Cli.getInstance();
    const spinnerSpinning = cli.spinner.isSpinning;

    /* c8 ignore next 3 */
    if (spinnerSpinning) {
      cli.spinner.stop();
    }

    const answer = await prompt.forSelection<T>(config);
    Cli.error('');

    // Restart the spinner if it was running before the prompt
    /* c8 ignore next 3 */
    if (spinnerSpinning) {
      cli.spinner.start();
    }

    return answer;
  }

  public static async promptForConfirmation(config: ConfirmationConfig): Promise<boolean> {
    const cli = Cli.getInstance();
    const spinnerSpinning = cli.spinner.isSpinning;

    /* c8 ignore next 3 */
    if (spinnerSpinning) {
      cli.spinner.stop();
    }

    const answer = await prompt.forConfirmation(config);
    Cli.error('');

    // Restart the spinner if it was running before the prompt
    /* c8 ignore next 3 */
    if (spinnerSpinning) {
      cli.spinner.start();
    }

    return answer;
  }

  public static async handleMultipleResultsFound<T>(message: string, values: { [key: string]: T }): Promise<T> {
    const prompt: boolean = Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.prompt, true);
    if (!prompt) {
      throw new Error(`${message} Found: ${Object.keys(values).join(', ')}.`);
    }

    Cli.error(`🌶️  ${message}`);
    const choices = Object.keys(values).map((choice: any) => { return { name: choice, value: choice }; });
    const response = await Cli.promptForSelection<string>({ message: `Please choose one:`, choices });

    return values[response];
  }

  private static removeShortOptions(args: { options: minimist.ParsedArgs }): { options: minimist.ParsedArgs } {
    const filteredArgs = JSON.parse(JSON.stringify(args));
    const optionsToRemove: string[] = Object.getOwnPropertyNames(args.options)
      .filter(option => option.length === 1 || option === '--');
    optionsToRemove.forEach(option => delete filteredArgs.options[option]);

    return filteredArgs;
  }

  private static loadOptionValuesFromFiles(args: { options: minimist.ParsedArgs }): void {
    const optionNames: string[] = Object.getOwnPropertyNames(args.options);
    optionNames.forEach(option => {
      const value = args.options[option];
      if (!value ||
        typeof value !== 'string' ||
        !value.startsWith('@')) {
        return;
      }

      const filePath: string = value.substr(1);
      // if the file doesn't exist, leave as-is, if it exists replace with
      // contents from the file
      if (fs.existsSync(filePath)) {
        args.options[option] = fs.readFileSync(filePath, 'utf-8');
      }
    });
  }

  public static shouldTrimOutput(output: string | undefined): boolean {
    return output === 'text';
  }
}
