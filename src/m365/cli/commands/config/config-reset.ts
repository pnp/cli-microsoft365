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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.key = args.options.key;
    return telemetryProps;
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-k, --key [key]',
        autocomplete: CliConfigResetCommand.optionNames
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.key) {
      if (CliConfigResetCommand.optionNames.indexOf(args.options.key) < 0) {
        return `${args.options.key} is not a valid setting. Allowed values: ${CliConfigResetCommand.optionNames.join(', ')}`;
      }
    }

    return true;
  }
}

module.exports = new CliConfigResetCommand();
