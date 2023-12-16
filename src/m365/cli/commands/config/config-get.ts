import { cli } from "../../../../cli/cli.js";
import { Logger } from "../../../../cli/Logger.js";
import GlobalOptions from "../../../../GlobalOptions.js";
import { settingsNames } from "../../../../settingsNames.js";
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from "../../commands.js";

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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await logger.log(cli.getConfig().get(args.options.key));
  }
}

export default new CliConfigGetCommand();