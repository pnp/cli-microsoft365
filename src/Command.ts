import type * as Chalk from 'chalk';
import appInsights from './appInsights';
import auth from './Auth';
import { Cli } from './cli';
import { Logger } from './cli/Logger';
import GlobalOptions from './GlobalOptions';
import request from './request';
import { GraphResponseError } from './utils';

export interface CommandOption {
  option: string;
  autocomplete?: string[]
}

export interface CommandHelp {
  (args: any, cbOrLog: (msg?: string) => void): void
}

export interface CommandTypes {
  string?: string[];
  boolean?: string[];
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
  protected _debug: boolean = false;
  protected _verbose: boolean = false;

  protected get debug(): boolean {
    return this._debug;
  }

  protected get verbose(): boolean {
    return this._verbose;
  }

  public abstract get name(): string;
  public abstract get description(): string;

  public abstract commandAction(logger: Logger, args: any, cb: () => void): void;

  protected showDeprecationWarning(logger: Logger, deprecated: string, recommended: string): void {
    const cli: Cli = Cli.getInstance();
    if (cli.currentCommandName &&
      cli.currentCommandName.indexOf(deprecated) === 0) {
      const chalk: typeof Chalk = require('chalk');
      logger.logToStderr(chalk.yellow(`Command '${deprecated}' is deprecated. Please use '${recommended}' instead`));
    }
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

  public action(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .restoreAuth()
      .then((): void => {
        this.initAction(args, logger);

        if (!auth.service.connected) {
          cb(new CommandError('Log in to Microsoft 365 first'));
          return;
        }

        this.commandAction(logger, args, cb);
      }, (error: any): void => {
        cb(new CommandError(error));
      });
  }

  public getTelemetryProperties(args: any): any {
    return {
      debug: this.debug.toString(),
      verbose: this.verbose.toString(),
      output: args.options.output,
      query: typeof args.options.query !== 'undefined'
    };
  }

  public alias(): string[] | undefined {
    return;
  }

  public autocomplete(): string[] | undefined {
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

  public options(): CommandOption[] {
    return [
      {
        option: '--query [query]'
      },
      {
        option: '-o, --output [output]',
        autocomplete: ['csv', 'json', 'text']
      },
      {
        option: '--verbose'
      },
      {
        option: '--debug'
      }
    ];
  }

  public optionSets(): string[][] | undefined {
    return;
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  public validate(args: any): boolean | string {
    return true;
  }

  public types(): CommandTypes | undefined {
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

  public getCommandName(): string {
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
    // console.log(JSON.stringify(response, null, 2));
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
    this._debug = args.options.debug || process.env.CLIMICROSOFT365_DEBUG === '1';
    this._verbose = this._debug || args.options.verbose || process.env.CLIMICROSOFT365_VERBOSE === '1';
    request.debug = this._debug;
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

    const knownOptions: CommandOption[] = this.options();
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
}
