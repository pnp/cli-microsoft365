import type * as Chalk from 'chalk';
import type * as Configstore from 'configstore';
import * as fs from 'fs';
import type { Inquirer } from 'inquirer';
import type * as JMESPath from 'jmespath';
import * as minimist from 'minimist';
import * as os from 'os';
import * as path from 'path';
import { Logger } from './Logger';
import Command, { CommandArgs, CommandError, CommandTypes } from '../Command';
import config from '../config';
import GlobalOptions from '../GlobalOptions';
import request from '../request';
import { settingsNames } from '../settingsNames';
import { formatting } from '../utils/formatting';
import { fsUtil } from '../utils/fsUtil';
import { md } from '../utils/md';
import { CommandInfo } from './CommandInfo';
import { CommandOptionInfo } from './CommandOptionInfo';
import { validation } from '../utils/validation';
import { telemetry } from '../telemetry';
const packageJSON = require('../../package.json');

export interface CommandOutput {
  stdout: string;
  stderr: string;
}

export class Cli {
  public commands: CommandInfo[] = [];
  /**
   * Command to execute
   */
  private commandToExecute: CommandInfo | undefined;
  /**
   * Name of the command specified through args
   */
  public currentCommandName: string | undefined;
  private optionsFromArgs: { options: minimist.ParsedArgs } | undefined;
  public commandsFolder: string = '';
  private static instance: Cli;
  private static defaultHelpMode = 'full';
  public static helpModes: string[] = ['options', 'examples', 'remarks', 'response', 'full'];

  private _config: Configstore | undefined;
  public get config(): Configstore {
    if (!this._config) {
      const configStore: typeof Configstore = require('configstore');
      this._config = new configStore(config.configstoreName);
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

  // eslint-disable-next-line @typescript-eslint/no-empty-function
  private constructor() {
  }

  public static getInstance(): Cli {
    if (!Cli.instance) {
      Cli.instance = new Cli();
    }

    return Cli.instance;
  }

  public async execute(commandsFolder: string, rawArgs: string[]): Promise<void> {
    this.commandsFolder = commandsFolder;

    // check if help for a specific command has been requested using the
    // 'm365 help xyz' format. If so, remove 'help' from the array of words
    // to use lazy loading commands but keep track of the fact that help should
    // be displayed
    let showHelp: boolean = false;
    if (rawArgs.length > 0 && rawArgs[0] === 'help') {
      showHelp = true;
      rawArgs.shift();
    }

    // parse args to see if a command has been specified and can be loaded
    // rather than loading all commands
    const parsedArgs: minimist.ParsedArgs = minimist(rawArgs);

    // load commands
    this.loadCommandFromArgs(parsedArgs._);

    if (this.currentCommandName) {
      for (let i = 0; i < this.commands.length; i++) {
        const command: CommandInfo = this.commands[i];
        if (command.name === this.currentCommandName ||
          (command.aliases &&
            command.aliases.indexOf(this.currentCommandName) > -1)) {
          this.commandToExecute = command;
          break;
        }
      }
    }

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
      this.printHelp(this.getHelpMode(parsedArgs));
      return Promise.resolve();
    }

    const optionsWithoutShorts = Cli.removeShortOptions(this.optionsFromArgs);

    try {
      // replace values staring with @ with file contents
      Cli.loadOptionValuesFromFiles(optionsWithoutShorts);
    }
    catch (e) {
      return this.closeWithError(e, optionsWithoutShorts);
    }

    try {
      // process options before passing them on to validation stage
      await this.commandToExecute.command.processOptions(optionsWithoutShorts.options);
    }
    catch (e: any) {
      return this.closeWithError(e.message, optionsWithoutShorts, false);
    }

    // if output not specified, set the configured output value (if any)
    if (optionsWithoutShorts.options.output === undefined) {
      optionsWithoutShorts.options.output = this.getSettingWithDefaultValue<string | undefined>(settingsNames.output, 'json');
    }

    const validationResult = await this.commandToExecute.command.validate(optionsWithoutShorts, this.commandToExecute);
    if (validationResult !== true) {
      return this.closeWithError(validationResult, optionsWithoutShorts, true);
    }

    return Cli
      .executeCommand(this.commandToExecute.command, optionsWithoutShorts)
      .then(_ => process.exit(0), err => this.closeWithError(err, optionsWithoutShorts));
  }

  public static async executeCommand(command: Command, args: { options: minimist.ParsedArgs }): Promise<void> {
    const logger: Logger = {
      log: (message: any): void => {
        const output: any = Cli.formatOutput(command, message, args.options);
        Cli.log(output);
      },
      logRaw: (message: any): void => Cli.log(message),
      logToStderr: (message: any): void => Cli.error(message)
    };

    if (args.options.debug) {
      logger.logToStderr(`Executing command ${command.name} with options ${JSON.stringify(args)}`);
    }

    // store the current command name, if any and set the name to the name of
    // the command to execute
    const cli = Cli.getInstance();
    const parentCommandName: string | undefined = cli.currentCommandName;
    cli.currentCommandName = command.getCommandName(cli.currentCommandName);

    try {
      await command.action(logger, args as any);

      if (args.options.debug || args.options.verbose) {
        const chalk: typeof Chalk = require('chalk');
        logger.logToStderr(chalk.green('DONE'));
      }
    }
    finally {
      // restore the original command name
      cli.currentCommandName = parentCommandName;
    }
  }

  public static async executeCommandWithOutput(command: Command, args: { options: minimist.ParsedArgs }, listener?: {
    stdout?: (message: any) => void,
    stderr?: (message: any) => void
  }): Promise<CommandOutput> {
    const log: string[] = [];
    const logErr: string[] = [];
    const logger: Logger = {
      log: (message: any): void => {
        const formattedMessage = Cli.formatOutput(command, message, args.options);
        if (listener && listener.stdout) {
          listener.stdout(formattedMessage);
        }
        log.push(formattedMessage);
      },
      logRaw: (message: any): void => {
        const formattedMessage = Cli.formatOutput(command, message, args.options);
        if (listener && listener.stdout) {
          listener.stdout(formattedMessage);
        }
        log.push(formattedMessage);
      },
      logToStderr: (message: any): void => {
        if (listener && listener.stderr) {
          listener.stderr(message);
        }
        logErr.push(message);
      }
    };

    if (args.options.debug) {
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

  public loadAllCommands(): void {
    const files: string[] = fsUtil.readdirR(this.commandsFolder) as string[];

    files.forEach(file => {
      if (file.indexOf(`${path.sep}commands${path.sep}`) > -1 &&
        file.indexOf(`${path.sep}assets${path.sep}`) < 0 &&
        file.endsWith('.js') &&
        !file.endsWith('.spec.js')) {
        try {
          const command: any = require(file);
          if (command instanceof Command) {
            this.loadCommand(command);
          }
        }
        catch (e) {
          this.closeWithError(e, { options: {} });
        }
      }
    });
  }

  /**
   * Loads command files into CLI based on the specified arguments.
   * 
   * @param commandNameWords Array of words specified as args
   */
  public loadCommandFromArgs(commandNameWords: string[]): void {
    this.currentCommandName = commandNameWords.join(' ');

    if (commandNameWords.length === 0) {
      this.loadAllCommands();
      return;
    }

    const isCompletionCommand: boolean = commandNameWords.indexOf('completion') > -1;
    if (isCompletionCommand) {
      this.loadAllCommands();
      return;
    }

    let commandFilePath = '';
    if (commandNameWords.length === 1) {
      commandFilePath = path.join(this.commandsFolder, 'commands', `${commandNameWords[0]}.js`);
    }
    else {
      if (commandNameWords.length === 2) {
        commandFilePath = path.join(this.commandsFolder, commandNameWords[0], 'commands', `${commandNameWords.join('-')}.js`);
      }
      else {
        commandFilePath = path.join(this.commandsFolder, commandNameWords[0], 'commands', commandNameWords[1], commandNameWords.slice(1).join('-') + '.js');
      }
    }

    this.loadCommandFromFile(commandFilePath);
  }

  /**
   * Loads command from the specified file into CLI. If can't find the file
   * or the file doesn't contain a command, loads all available commands.
   * 
   * @param commandFilePath File path of the file with command to load
   */
  private loadCommandFromFile(commandFilePath: string): void {
    if (!fs.existsSync(commandFilePath)) {
      this.loadAllCommands();
      return;
    }

    try {
      const command: any = require(commandFilePath);
      if (command instanceof Command) {
        this.loadCommand(command);
      }
      else {
        this.loadAllCommands();
      }
    }
    catch {
      this.loadAllCommands();
    }
  }

  public static getCommandInfo(command: Command): CommandInfo {
    return {
      aliases: command.alias(),
      name: command.name,
      command: command,
      options: this.getCommandOptions(command),
      defaultProperties: command.defaultProperties()
    };
  }

  private loadCommand(command: Command): void {
    this.commands.push(Cli.getCommandInfo(command));
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

  private static formatOutput(command: Command, logStatement: any, options: GlobalOptions): any {
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
      const jmespath: typeof JMESPath = require('jmespath');
      try {
        logStatement = jmespath.search(logStatement, options.query);
      }
      catch (e: any) {
        const message = `JMESPath query error. ${e.message}. See https://jmespath.org/specification.html for more information`;
        Cli.getInstance().closeWithError(message, { options }, false);
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
      const chalk: typeof Chalk = require('chalk');
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
    if (options.output === 'text' || options.output === 'csv') {
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
        return command.getCsvOutput(logStatement);
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
      Cli.log(`CLI for Microsoft 365 v${packageJSON.version}`);
      Cli.log(`${packageJSON.description}`);
      Cli.log();

      properties.command = 'commandList';
      this.printAvailableCommands();
    }

    telemetry.trackEvent('help', properties);

    process.exit(exitCode);
  }

  private printCommandHelp(helpMode: string): void {
    let helpFilePath = '';
    let commandNameWords: string[] = [];
    if (this.commandToExecute) {
      commandNameWords = (this.commandToExecute.name).split(' ');
    }
    const pathChunks: string[] = [this.commandsFolder, '..', '..', 'docs', 'docs', 'cmd'];

    if (commandNameWords.length === 1) {
      pathChunks.push(`${commandNameWords[0]}.md`);
    }
    else {
      if (commandNameWords.length === 2) {
        pathChunks.push(commandNameWords[0], `${commandNameWords.join('-')}.md`);
      }
      else {
        pathChunks.push(commandNameWords[0], commandNameWords[1], commandNameWords.slice(1).join('-') + '.md');
      }
    }

    helpFilePath = path.join(...pathChunks);

    if (fs.existsSync(helpFilePath)) {
      let helpContents = fs.readFileSync(helpFilePath, 'utf8');
      helpContents = this.getHelpSection(helpMode, helpContents);
      helpContents = md.md2plain(helpContents, path.join(this.commandsFolder, '..', '..', 'docs'));
      Cli.log();
      Cli.log(helpContents);
    }
  }

  private getHelpMode(options: any): string {
    const h = options.h;
    const help = options.help;

    // user passed -h or --help, let's see if they passed a specific mode
    // or requested the default
    if (h || help) {
      const helpMode: boolean | string = h ?? help;

      if (typeof helpMode === 'boolean' || typeof helpMode !== 'string') {
        // requested default mode or passed a number, let's use default
        return this.getSettingWithDefaultValue<string>(settingsNames.helpMode, Cli.defaultHelpMode);
      }
      else {
        const lowerCaseHelpMode = helpMode.toLowerCase();

        if (Cli.helpModes.indexOf(lowerCaseHelpMode) < 0) {
          Cli.getInstance().closeWithError(`Unknown help mode ${helpMode}. Allowed values are ${Cli.helpModes.join(', ')}`, { options }, false);
          return ''; // noop
        }
        else {
          return lowerCaseHelpMode;
        }
      }
    }
    else {
      return this.getSettingWithDefaultValue<string>(settingsNames.helpMode, Cli.defaultHelpMode);
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

      for (const commandName in commandsToPrint) {
        Cli.log(`  ${`${commandName} [options]`.padEnd(maxLength, ' ')}  ${commandsToPrint[commandName].command.description}`);
      }
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

  private closeWithError(error: any, args: CommandArgs, showHelpIfEnabled: boolean = false): void {
    const chalk: typeof Chalk = require('chalk');
    let exitCode: number = 1;

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

    Cli.error(errorMessage);

    if (showHelpIfEnabled &&
      this.getSettingWithDefaultValue<boolean>(settingsNames.showHelpOnFailure, showHelpIfEnabled)) {
      this.printHelp(this.getHelpMode(args.options), exitCode);
    }
    else {
      process.exit(exitCode);
    }

    // will never be run. Required for testing where we're stubbing process.exit
    /* c8 ignore next */
    throw new Error();
    /* c8 ignore next */
  }

  public static log(message?: any, ...optionalParams: any[]): void {
    if (message) {
      console.log(message, ...optionalParams);
    }
    else {
      console.log();
    }
  }

  private static error(message?: any, ...optionalParams: any[]): void {
    const errorOutput: string = Cli.getInstance().getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');
    if (errorOutput === 'stdout') {
      console.log(message, ...optionalParams);
    }
    else {
      console.error(message, ...optionalParams);
    }
  }

  public static async prompt<T>(options: any): Promise<T> {
    const inquirer: Inquirer = require('inquirer');
    return await inquirer.prompt(options) as T;
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
}