import appInsights from './appInsights';
import GlobalOptions from './GlobalOptions';

export interface CommandOption {
  option: string;
  description: string;
  autocomplete?: string[]
}

export interface CommandAction {
  (this: CommandInstance, args: any, cb: () => void): void
}

export interface CommandValidate {
  (args: any): boolean | string
}

export interface CommandHelp {
  (args: any, log: (help: string) => void): void
}

export interface CommandCancel {
  (): void
}

export interface CommandTypes {
  string?: string[];
  boolean?: string[];
}

export class CommandError {
  constructor(public message: string) {
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

interface CommandArgs {
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

  public abstract commandAction(cmd: CommandInstance, args: any, cb: () => void): void;

  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: () => void) {
      cmd._debug = args.options.debug || false;
      cmd._verbose = cmd._debug || args.options.verbose || false;

      appInsights.trackEvent({
        name: cmd.getCommandName(),
        properties: cmd.getTelemetryProperties(args)
      });

      cmd.commandAction(this, args, cb);
    }
  }

  public getTelemetryProperties(args: any): any {
    return {
      debug: this.debug.toString(),
      verbose: this.verbose.toString()
    };
  }

  public autocomplete(): string[] | undefined {
    return;
  }

  public options(): CommandOption[] {
    return [
      {
        option: '-o, --output <output>',
        description: 'Output type. json|text. Default text',
        autocomplete: ['json', 'text']
      },
      {
        option: '--verbose',
        description: 'Runs command with verbose logging'
      },
      {
        option: '--debug',
        description: 'Runs command with debug logging'
      }
    ];
  }

  public help(): CommandHelp | undefined {
    return;
  }

  public validate(): CommandValidate | undefined {
    return;
  }

  public cancel(): CommandCancel | undefined {
    return;
  }

  public types(): CommandTypes | undefined {
    return;
  }

  public init(vorpal: Vorpal): void {
    const cmd: VorpalCommand = vorpal
      .command(this.name, this.description, this.autocomplete())
      .action(this.action());
    const options: CommandOption[] = this.options();
    options.forEach((o: CommandOption): void => {
      cmd.option(o.option, o.description, o.autocomplete);
    });
    const validate: CommandValidate | undefined = this.validate();
    if (validate) {
      cmd.validate(validate);
    }
    const cancel: CommandCancel | undefined = this.cancel();
    if (cancel) {
      cmd.cancel(cancel);
    }
    const help: CommandHelp | undefined = this.help();
    if (help) {
      cmd.help(help);
    }
    const types: CommandTypes | undefined = this.types();
    if (types) {
      cmd.types(types);
    }
  }

  public getCommandName(): string {
    let commandName: string = this.name;
    let pos: number = commandName.indexOf('<');
    let pos1: number = commandName.indexOf('[');
    if (pos > -1 || pos1 > -1) {
      if (pos1 > -1) {
        pos = pos1;
      }

      commandName = commandName.substr(0, pos).trim();
    }

    return commandName;
  }

  protected handleRejectedODataPromise(rawResponse: any, cmd: CommandInstance, callback: () => void): void {
    if (rawResponse instanceof Error) {
      cmd.log(new CommandError(rawResponse.message));
      callback();
      return;
    }

    const res: any = JSON.parse(JSON.stringify(rawResponse));
    if (res.error) {
      try {
        const err: ODataError = JSON.parse(res.error);
        cmd.log(new CommandError(err['odata.error'].message.value));
      }
      catch {
        cmd.log(new CommandError(res.error));
      }
    }
    else {
      cmd.log(new CommandError(rawResponse));
    }

    callback();
  }

  protected handleError(rawResponse: any, cmd: CommandInstance): void {
    if (rawResponse instanceof Error) {
      cmd.log(new CommandError(rawResponse.message));
    }
    else {
      cmd.log(new CommandError(rawResponse));          
    }
  }

  protected handleRejectedPromise(rawResponse: any, cmd: CommandInstance, callback: () => void): void {
    this.handleError(rawResponse, cmd);

    callback();
  }
}