import { Cli, Logger } from "../../../../cli";
import GlobalOptions from "../../../../GlobalOptions";
import { settingsNames } from "../../../../settingsNames";
import AnonymousCommand from "../../../base/AnonymousCommand";
import commands from "../../commands";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  key: string;
}

class CliConfigGetCommand extends AnonymousCommand {
  private static readonly optionNames: string[] = Object.getOwnPropertyNames(settingsNames);

  public get name(): string {
    return commands.CONFIG_GET;
  }

  public get description(): string {
    return 'Gets value of a CLI for Microsoft 365 configuration option';
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
        option: '-k, --key <key>',
        autocomplete: CliConfigGetCommand.optionNames
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (CliConfigGetCommand.optionNames.indexOf(args.options.key) < 0) {
          return `${args.options.key} is not a valid setting. Allowed values: ${CliConfigGetCommand.optionNames.join(', ')}`;
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    logger.log(Cli.getInstance().config.get(args.options.key));
    cb();
  }
}

module.exports = new CliConfigGetCommand();