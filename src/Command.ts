import type * as Chalk from 'chalk';
import type { Inquirer } from 'inquirer';
import * as os from 'os';
import appInsights from './appInsights';
import auth from './Auth';
import { Cli, CommandInfo, CommandOptionInfo } from './cli';
import { Logger } from './cli/Logger';
import GlobalOptions from './GlobalOptions';
import request from './request';
import { settingsNames } from './settingsNames';
import { accessToken, GraphResponseError } from './utils';

export interface CommandOption {
  option: string;
  autocomplete?: string[]
}

export interface CommandHelp {
  (args: any, cbOrLog: (msg?: string) => void): void
}

export interface CommandTypes {
  string: string[];
  boolean: string[];
}

export class CommandError {
  constructor(public message: string, public code?: number) {
  }
}

export class CommandErrorWithOutput {
  constructor(public error: CommandError, public stderr?: string) {
  }
}

export interface ODataError {
  "odata.error": {
    code: string;
    message: {
      lang: string;
      value: string;
    }
  }
}

export interface CommandArgs {
  options: GlobalOptions;
}

export default abstract class Command {
  protected debug: boolean = false;
  protected verbose: boolean = false;

  public telemetry: ((args: any) => void)[] = [];
  protected telemetryProperties: any = {};

  public options: CommandOption[] = [];
  public optionSets: string[][] = [];
  public types: CommandTypes = {
    boolean: [],
    string: []
  };

  protected validators: ((args: any, command: CommandInfo) => Promise<boolean | string>)[] = [];

  public abstract get name(): string;
  public abstract get description(): string;

  constructor() {
    // These functions must be defined with # so that they're truly private
    // otherwise you'll get a ts2415 error (Types have separate declarations of
    // a private property 'x'.).
    // `private` in TS is a design-time flag and private members end-up being
    // regular class properties that would collide on runtime, which is why we
    // need the extra `#`

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        debug: this.debug.toString(),
        verbose: this.verbose.toString(),
        output: args.options.output,
        query: typeof args.options.query !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--query [query]' },
      {
        option: '-o, --output [output]',
        autocomplete: ['csv', 'json', 'text']
      },
      { option: '--verbose' },
      { option: '--debug' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      (args, command) => this.validateUnknownOptions(args, command),
      (args, command) => this.validateRequiredOptions(args, command),
      (args, command) => this.validateOptionSets(args, command),
    );
  }

  private async validateUnknownOptions(args: CommandArgs, command: CommandInfo): Promise<string | boolean> {
    if (this.allowUnknownOptions()) {
      return true;
    }

    // if the command doesn't allow unknown options, check if all specified
    // options match command options
    for (const optionFromArgs in args.options) {
      let matches: boolean = false;

      for (let i = 0; i < command.options.length; i++) {
        const option: CommandOptionInfo = command.options[i];
        if (optionFromArgs === option.long ||
          optionFromArgs === option.short) {
          matches = true;
          break;
        }
      }

      if (!matches) {
        return `Invalid option: '${optionFromArgs}'${os.EOL}`;
      }
    }

    return true;
  }

  private async validateRequiredOptions(args: CommandArgs, command: CommandInfo): Promise<string | boolean> {
    const shouldPrompt = Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.prompt, false);

    let inquirer: Inquirer | undefined;
    let prompted: boolean = false;
    for (let i = 0; i < command.options.length; i++) {
      if (!command.options[i].required ||
        typeof args.options[command.options[i].name] !== 'undefined') {
        continue;
      }

      if (!shouldPrompt) {
        return `Required option ${command.options[i].name} not specified`;
      }

      if (!prompted) {
        prompted = true;
        Cli.log('Provide values for the following parameters:');
      }

      if (!inquirer) {
        inquirer = require('inquirer');
      }

      const missingRequireOptionValue = await (inquirer as Inquirer)
        .prompt({
          name: 'missingRequireOptionValue',
          message: `${command.options[i].name}: `
        })
        .then(result => result.missingRequireOptionValue);

      args.options[command.options[i].name] = missingRequireOptionValue;
    }

    return true;
  }

  private async validateOptionSets(args: CommandArgs, command: CommandInfo): Promise<string | boolean> {
    const optionsSets: string[][] | undefined = command.command.optionSets;
    if (!optionsSets || optionsSets.length === 0) {
      return true;
    }

    const argsOptions: string[] = Object.keys(args.options);
    for (const optionSet of optionsSets) {
      const commonOptions = argsOptions.filter(opt => optionSet.includes(opt));

      if (commonOptions.length === 0) {
        return `Specify one of the following options: ${optionSet.map(opt => opt).join(', ')}.`;
      }

      if (commonOptions.length > 1) {
        return `Specify one of the following options: ${optionSet.map(opt => opt).join(', ')}, but not multiple.`;
      }
    }

    return true;
  }

  public alias(): string[] | undefined {
    return;
  }

  /**
   * Returns list of properties that should be returned in the text output.
   * Returns all properties if no default properties specified
   */
  public defaultProperties(): string[] | undefined {
    return;
  }

  public allowUnknownOptions(): boolean | undefined {
    return;
  }

  /**
   * Processes options after resolving them from the user input and before
   * passing them on to command action for execution. Used for example for
   * expanding server-relative URLs to absolute in spo commands
   * @param options Object that contains command's options
   */
  // eslint-disable-next-line @typescript-eslint/no-unused-vars, @typescript-eslint/no-empty-function
  public async processOptions(options: any): Promise<void> {
  }

  public abstract commandAction(logger: Logger, args: any, cb: () => void): void;

  public action(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .restoreAuth()
      .then((): void => {
        this.initAction(args, logger);

        if (!auth.service.connected) {
          cb(new CommandError('Log in to Microsoft 365 first'));
          return;
        }

        try {
          this.loadValuesFromAccessToken(args);
          this.commandAction(logger, args, cb);
        }
        catch (ex) {
          cb(new CommandError(ex as any));
        }
      }, (error: any): void => {
        cb(new CommandError(error));
      });
  }

  public async validate(args: CommandArgs, command: CommandInfo): Promise<boolean | string> {
    for (const validate of this.validators) {
      const result = await validate(args, command);
      if (result !== true) {
        return result;
      }
    }

    return true;
  }

  public getCommandName(alias?: string): string {
    if (alias &&
      this.alias()?.includes(alias)) {
      return alias;
    }

    let commandName: string = this.name;
    let pos: number = commandName.indexOf('<');
    const pos1: number = commandName.indexOf('[');
    if (pos > -1 || pos1 > -1) {
      if (pos1 > -1) {
        pos = pos1;
      }

      commandName = commandName.substr(0, pos).trim();
    }

    return commandName;
  }

  protected handleRejectedODataPromise(rawResponse: any, logger: Logger, callback: (err?: any) => void): void {
    const res: any = JSON.parse(JSON.stringify(rawResponse));
    if (res.error) {
      try {
        const err: ODataError = JSON.parse(res.error);
        callback(new CommandError(err['odata.error'].message.value));
      }
      catch {
        try {
          const graphResponseError: GraphResponseError = res.error;
          if (graphResponseError.error.code) {
            callback(new CommandError(graphResponseError.error.code + " - " + graphResponseError.error.message));
          }
          else {
            callback(new CommandError(graphResponseError.error.message));
          }
        }
        catch {
          callback(new CommandError(res.error));
        }
      }
    }
    else {
      if (rawResponse instanceof Error) {
        callback(new CommandError(rawResponse.message));
      }
      else {
        callback(new CommandError(rawResponse));
      }
    }
  }

  protected handleRejectedODataJsonPromise(response: any, logger: Logger, callback: (err?: any) => void): void {
    if (response.error &&
      response.error['odata.error'] &&
      response.error['odata.error'].message) {
      return callback(new CommandError(response.error['odata.error'].message.value));
    }

    if (!response.error) {
      if (response instanceof Error) {
        return callback(new CommandError(response.message));
      }
      else {
        return callback(new CommandError(response));
      }
    }

    if (response.error.error &&
      response.error.error.message) {
      return callback(new CommandError(response.error.error.message));
    }

    if (response.error.message) {
      return callback(new CommandError(response.error.message));
    }

    if (response.error.error_description) {
      return callback(new CommandError(response.error.error_description));
    }

    try {
      const error: any = JSON.parse(response.error);
      if (error &&
        error.error &&
        error.error.message) {
        callback(new CommandError(error.error.message));
      }
      else {
        callback(new CommandError(response.error));
      }
    }
    catch {
      callback(new CommandError(response.error));
    }
  }

  protected handleError(rawResponse: any, logger: Logger, callback: (err?: any) => void): void {
    if (rawResponse instanceof Error) {
      callback(new CommandError(rawResponse.message));
    }
    else {
      callback(new CommandError(rawResponse));
    }
  }

  protected handleRejectedPromise(rawResponse: any, logger: Logger, callback: (err?: any) => void): void {
    this.handleError(rawResponse, logger, callback);
  }

  protected initAction(args: CommandArgs, logger: Logger): void {
    this.debug = args.options.debug || process.env.CLIMICROSOFT365_DEBUG === '1';
    this.verbose = this.debug || args.options.verbose || process.env.CLIMICROSOFT365_VERBOSE === '1';
    request.debug = this.debug;
    request.logger = logger;

    appInsights.trackEvent({
      name: this.getUsedCommandName(),
      properties: this.getTelemetryProperties(args)
    });
    appInsights.flush();
  }

  protected getUnknownOptions(options: any): any {
    const unknownOptions: any = JSON.parse(JSON.stringify(options));
    // remove minimist catch-all option
    delete unknownOptions._;

    const knownOptions: CommandOption[] = this.options;
    const longOptionRegex: RegExp = /--([^\s]+)/;
    const shortOptionRegex: RegExp = /-([a-z])\b/;
    knownOptions.forEach(o => {
      const longOptionName: string = (longOptionRegex.exec(o.option) as RegExpExecArray)[1];
      delete unknownOptions[longOptionName];

      // short names are optional so we need to check if the current command has
      // one before continuing
      const shortOptionMatch: RegExpExecArray | null = shortOptionRegex.exec(o.option);
      if (shortOptionMatch) {
        const shortOptionName: string = shortOptionMatch[1];
        delete unknownOptions[shortOptionName];
      }
    });

    return unknownOptions;
  }

  protected trackUnknownOptions(telemetryProps: any, options: any): void {
    const unknownOptions: any = this.getUnknownOptions(options);
    const unknownOptionsNames: string[] = Object.getOwnPropertyNames(unknownOptions);
    unknownOptionsNames.forEach(o => {
      telemetryProps[o] = true;
    });
  }

  protected addUnknownOptionsToPayload(payload: any, options: any): void {
    const unknownOptions: any = this.getUnknownOptions(options);
    const unknownOptionsNames: string[] = Object.getOwnPropertyNames(unknownOptions);
    unknownOptionsNames.forEach(o => {
      payload[o] = unknownOptions[o];
    });
  }

  private loadValuesFromAccessToken(args: CommandArgs) {
    if (!auth.service.accessTokens[auth.defaultResource]) {
      return;
    }

    const token = auth.service.accessTokens[auth.defaultResource].accessToken;
    const optionNames: string[] = Object.getOwnPropertyNames(args.options);
    optionNames.forEach(option => {
      const value = args.options[option];
      if (!value || typeof value !== 'string') {
        return;
      }

      const lowerCaseValue = value.toLowerCase();
      if (lowerCaseValue === '@meid') {
        args.options[option] = accessToken.getUserIdFromAccessToken(token);
      }
      if (lowerCaseValue === '@meusername') {
        args.options[option] = accessToken.getUserNameFromAccessToken(token);
      }
    });
  }

  protected showDeprecationWarning(logger: Logger, deprecated: string, recommended: string): void {
    const cli: Cli = Cli.getInstance();
    if (cli.currentCommandName &&
      cli.currentCommandName.indexOf(deprecated) === 0) {
      const chalk: typeof Chalk = require('chalk');
      logger.logToStderr(chalk.yellow(`Command '${deprecated}' is deprecated. Please use '${recommended}' instead`));
    }
  }

  protected warn(logger: Logger, warning: string): void {
    const chalk: typeof Chalk = require('chalk');
    logger.logToStderr(chalk.yellow(warning));
  }

  protected getUsedCommandName(): string {
    const cli: Cli = Cli.getInstance();
    const commandName: string = this.getCommandName();
    if (!cli.currentCommandName) {
      return commandName;
    }

    if (cli.currentCommandName &&
      cli.currentCommandName.indexOf(commandName) === 0) {
      return commandName;
    }

    // since the command was called by something else than its name
    // it must have aliases
    const aliases: string[] = this.alias() as string[];

    for (let i: number = 0; i < aliases.length; i++) {
      if (cli.currentCommandName.indexOf(aliases[i]) === 0) {
        return aliases[i];
      }
    }

    // shouldn't happen because the command is called either by its name or alias
    return '';
  }

  private getTelemetryProperties(args: any): any {
    this.telemetry.forEach(t => t(args));
    return this.telemetryProperties;
  }
}
