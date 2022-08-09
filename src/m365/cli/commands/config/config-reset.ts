import { Cli, Logger } from "../../../../cli";
import GlobalOptions from "../../../../GlobalOptions";
import { settingsNames } from "../../../../settingsNames";
import AnonymousCommand from "../../../base/AnonymousCommand";
import commands from "../../commands";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  key?: string;
}

class CliConfigResetCommand extends AnonymousCommand {
  private static readonly optionNames: string[] = Object.getOwnPropertyNames(settingsNames);

  public get name(): string {
    return commands.CONFIG_RESET;
  }

  public get description(): string {
    return 'Resets the specified CLI configuration option to its default value';
  }

  constructor() {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        key: args.options.key
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-k, --key [key]',
        autocomplete: CliConfigResetCommand.optionNames
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.key) {
          if (CliConfigResetCommand.optionNames.indexOf(args.options.key) < 0) {
            return `${args.options.key} is not a valid setting. Allowed values: ${CliConfigResetCommand.optionNames.join(', ')}`;
          }
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (args.options.key) {
      Cli.getInstance().config.delete(args.options.key);
    }
    else {
      Cli.getInstance().config.clear();
    }

    cb();
  }
}

module.exports = new CliConfigResetCommand();
