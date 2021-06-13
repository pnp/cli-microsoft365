import { Cli, Logger } from "../../../../cli";
import { CommandOption } from "../../../../Command";
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.key = args.options.key;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    logger.log(Cli.getInstance().config.get(args.options.key));
    cb();
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-k, --key <key>',
        autocomplete: CliConfigGetCommand.optionNames
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (CliConfigGetCommand.optionNames.indexOf(args.options.key) < 0) {
      return `${args.options.key} is not a valid setting. Allowed values: ${CliConfigGetCommand.optionNames.join(', ')}`;
    }

    return true;
  }
}

module.exports = new CliConfigGetCommand();