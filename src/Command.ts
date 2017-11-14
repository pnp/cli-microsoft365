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

export default abstract class Command {
  public abstract get name(): string;
  public abstract get description(): string;
  public abstract get action(): CommandAction;

  public autocomplete(): string[] | undefined {
    return;
  }

  public options(): CommandOption[] | undefined {
    return [{
      option: '--verbose',
      description: 'Runs command with verbose logging'
    }];
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
      .action(this.action);
    const options: CommandOption[] | undefined = this.options();
    if (options) {
      options.forEach((o: CommandOption): void => {
        cmd.option(o.option, o.description, o.autocomplete);
      });
    }
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
}