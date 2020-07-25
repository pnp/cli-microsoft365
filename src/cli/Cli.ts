import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as minimist from 'minimist';
import * as jmespath from 'jmespath';
import * as chalk from 'chalk';
import * as inquirer from 'inquirer';
import * as markshell from 'markshell';
import Table = require('easy-table');
import Command, { CommandError, CommandValidate } from '../Command';
import { CommandInfo } from './CommandInfo';
import { CommandOptionInfo } from './CommandOptionInfo';
const packageJSON = require('../../package.json');

export class Cli {
  public commands: CommandInfo[] = [];
  /**
   * Command to execute
   */
  private commandToExecute: CommandInfo | undefined;
  /**
   * Name of the command specified through args
   */
  private currentCommandName: string | undefined;
  private optionsFromArgs: { options: minimist.ParsedArgs } | undefined;
  private commandsFolder: string = '';
  private static instance: Cli;

  private constructor() {
  }

  public static getInstance(): Cli {
    if (!Cli.instance) {
      Cli.instance = new Cli();
    }

    return Cli.instance;
  }

  public execute(commandsFolder: string, rawArgs: string[]): Promise<void> {
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
      this.optionsFromArgs = {
        options: this.getCommandOptionsFromArgs(rawArgs, this.commandToExecute)
      };
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
      this.printHelp();
      return Promise.resolve();
    }

    // if the command doesn't allow unknown options, check if all specified
    // options match command options
    if (!this.commandToExecute.command.allowUnknownOptions()) {
      for (let optionFromArgs in this.optionsFromArgs.options) {
        if (optionFromArgs === '_') {
          // ignore catch-all option
          continue;
        }

        let matches: boolean = false;

        for (let i = 0; i < this.commandToExecute.options.length; i++) {
          const option: CommandOptionInfo = this.commandToExecute.options[i];
          if (optionFromArgs === option.long ||
            optionFromArgs === option.short) {
            matches = true;
            break;
          }
        }

        if (!matches) {
          return this.closeWithError(`Invalid option: '${optionFromArgs}'${os.EOL}`, true);
        }
      }
    }

    // validate options
    // validate required options
    for (let i = 0; i < this.commandToExecute.options.length; i++) {
      if (this.commandToExecute.options[i].required &&
        typeof this.optionsFromArgs.options[this.commandToExecute.options[i].name] === 'undefined') {
        return this.closeWithError(`Required option ${this.commandToExecute.options[i].name} not specified`, true);
      }
    }

    // validate using command's validate method if specified
    const validate: CommandValidate | undefined = this.commandToExecute.command.validate();
    if (validate) {
      const validationResult: boolean | string = validate(this.optionsFromArgs);
      if (typeof validationResult === 'string') {
        return this.closeWithError(validationResult, true);
      }
    }

    return Cli
      .executeCommand(this.commandToExecute.name, this.commandToExecute.command, this.optionsFromArgs)
      .then(_ => process.exit(0), err => this.closeWithError(err));
  }

  public static executeCommand(commandName: string, command: Command, args: { options: minimist.ParsedArgs }): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const commandInstance = {
        commandWrapper: {
          command: commandName
        },
        action: command.action(),
        log: (message: any): void => {
          const output: any = Cli.logOutput(message, args.options);
          Cli.log(output);
        },
        prompt: (options: any, cb: (result: any) => void) => {
          inquirer
            .prompt(options)
            .then(result => cb(result));
        }
      };

      if (args.options.debug) {
        commandInstance.log(`Executing command ${command.name} with options ${JSON.stringify(args)}`);
      }

      commandInstance.action(args, (err: any): void => {
        if (err) {
          return reject(err);
        }

        resolve();
      });
    });
  }

  public static executeCommandWithOutput(commandName: string, command: Command, args: { options: minimist.ParsedArgs }): Promise<string> {
    return new Promise((resolve: (result: string) => void, reject: (error: any) => void): void => {
      const log: string[] = [];
      const commandInstance = {
        commandWrapper: {
          command: commandName
        },
        action: command.action(),
        log: (message: any): void => {
          log.push(message);
        },
        prompt: (options: any, cb: (result: any) => void) => {
          inquirer
            .prompt(options)
            .then(result => cb(result));
        }
      };

      if (args.options.debug) {
        Cli.log(`Executing command ${command.name} with options ${JSON.stringify(args)}`);
      }

      commandInstance.action(args, (err: any): void => {
        if (err) {
          return reject(err);
        }

        resolve(log.join());
      });
    });
  }

  public loadAllCommands(): void {
    const files: string[] = Cli.readdirR(this.commandsFolder) as string[];

    files.forEach(file => {
      if (file.indexOf(`${path.sep}commands${path.sep}`) > -1 &&
        file.endsWith('.js') &&
        !file.endsWith('.spec.js')) {
        try {
          const command: any = require(file);
          if (command instanceof Command) {
            this.loadCommand(command);
          }
        }
        catch (e) {
          this.closeWithError(e);
        }
      }
    });
  };

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

  private static readdirR(dir: string): string | string[] {
    return fs.statSync(dir).isDirectory()
      ? Array.prototype.concat(...fs.readdirSync(dir).map(f => Cli.readdirR(path.join(dir, f))))
      : dir;
  };

  private loadCommand(command: Command): void {
    this.commands.push({
      aliases: command.alias(),
      name: command.name,
      command: command,
      options: this.getCommandOptions(command)
    });
  }

  private getCommandOptions(command: Command): CommandOptionInfo[] {
    const options: CommandOptionInfo[] = [];

    command.options().forEach(option => {
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

    if (commandInfo) {
      const commandTypes = commandInfo.command.types();
      if (commandTypes) {
        minimistOptions.string = commandTypes.string;
        minimistOptions.boolean = commandTypes.boolean;
      }
      minimistOptions.alias = {};
      commandInfo.options.forEach(option => {
        if (option.short && option.long) {
          (minimistOptions.alias as any)[option.short] = option.long;
        }
      });
    }

    return minimist(args, minimistOptions);
  }

  private static logOutput(logStatement: any, options: any): any {
    if (logStatement instanceof Date) {
      return logStatement.toString();
    }

    const logStatementType: string = typeof logStatement;

    if (logStatementType === 'undefined') {
      return logStatement;
    }

    if (options.query &&
      !options.help) {
      logStatement = jmespath.search(logStatement, options.query);
    }

    if (options.output === 'json') {
      return JSON.stringify(logStatement, null, 2);
    }

    if (logStatement instanceof CommandError) {
      return chalk.red(`Error: ${logStatement.message}`);
    }

    let arrayType: string = '';
    if (!Array.isArray(logStatement)) {
      logStatement = [logStatement];
      arrayType = logStatementType;
    }
    else {
      for (let i: number = 0; i < logStatement.length; i++) {
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

    if (logStatement.length === 1) {
      const obj: any = logStatement[0];
      const propertyNames: string[] = [];
      Object.getOwnPropertyNames(obj).forEach(p => {
        propertyNames.push(p);
      });

      let longestPropertyLength: number = 0;
      propertyNames.forEach(p => {
        if (p.length > longestPropertyLength) {
          longestPropertyLength = p.length;
        }
      });

      const output: string[] = [];
      propertyNames.sort().forEach(p => {
        output.push(`${p.length < longestPropertyLength ? p + new Array(longestPropertyLength - p.length + 1).join(' ') : p}: ${Array.isArray(obj[p]) || typeof obj[p] === 'object' ? JSON.stringify(obj[p]) : obj[p]}`);
      });

      return output.join('\n') + '\n';
    }
    else {
      const t: Table = new Table();
      logStatement.forEach((r: any) => {
        if (typeof r !== 'object') {
          return;
        }

        Object.getOwnPropertyNames(r).forEach(p => {
          t.cell(p, r[p]);
        });
        t.newRow();
      });

      return t.toString();
    }
  }

  private printHelp(exitCode: number = 0, std: any = Cli.log): void {
    if (this.commandToExecute) {
      this.printCommandHelp();
    }
    else {
      Cli.log();
      Cli.log(`CLI for Microsoft 365 v${packageJSON.version}`);
      Cli.log(`${packageJSON.description}`);
      Cli.log();

      this.printAvailableCommands();
    }

    process.exit(exitCode);
  }

  private printCommandHelp(): void {
    let helpFilePath = '';
    const commandNameWords = (this.optionsFromArgs as { options: minimist.ParsedArgs }).options._;
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
      Cli.log();

      const theme = markshell.getTheme();
      const admonitionStyles = theme.admonitions.getStyles();
      admonitionStyles.indent.beforeIndent = 0;
      admonitionStyles.indent.titleIndent = 3;
      admonitionStyles.indent.afterIndent = 0;
      theme.admonitions.setStyles(admonitionStyles);
      theme.indents.definitionList = 2;
      theme.headline = chalk.white;
      theme.inlineCode = chalk.cyan;
      theme.sourceCodeTheme = 'solarizelight';
      markshell.setTheme(theme);

      Cli.log(markshell.toRawContent(helpFilePath));
    }
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
    }

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

      for (let commandName in commandsToPrint) {
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

      for (let commandGroup in commandGroupsToPrint) {
        Cli.log(`  ${`${commandGroup} *`.padEnd(maxLength, ' ')}  ${commandGroupsToPrint[commandGroup]} command${commandGroupsToPrint[commandGroup] === 1 ? '' : 's'}`);
      }
    }

    Cli.log();
  }

  private closeWithError(error: any, showHelp: boolean = false): Promise<void> {
    let exitCode: number = 1;

    if (error instanceof CommandError) {
      if (error.code) {
        exitCode = error.code;
      }

      Cli.error(chalk.red(`Error: ${error.message}`));
    }
    else {
      Cli.error(chalk.red(`Error: ${error}`));
    }

    if (showHelp) {
      Cli.log();
      this.printHelp(exitCode);
    }
    else {
      process.exit(exitCode);
    }

    // will never be run. Required for testing where we're stubbing process.exit
    /* c8 ignore next */
    return Promise.reject();
    /* c8 ignore next */
  }

  private static log(message?: any, ...optionalParams: any[]): void {
    if (message) {
      console.log(message, ...optionalParams);
    }
    else {
      console.log();
    }
  }

  private static error(message?: any, ...optionalParams: any[]): void {
    console.error(message, ...optionalParams);
  }
}