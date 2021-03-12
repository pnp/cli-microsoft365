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
  value: string;
}

class CliConfigSetCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CONFIG_SET;
  }

  public get description(): string {
    return 'Manage global configuration settings about the CLI for Microsoft 365';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps[args.options.key] = args.options.value;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let value: any = undefined;

    switch (args.options.key) {
      case settingsNames.showHelpOnFailure:
        value = args.options.value === "true";
        break;
    }

    console.log(Cli.getInstance().config.path);
    Cli.getInstance().config.set(args.options.key, value);
    cb();
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-k, --key <key>',
        autocomplete: [settingsNames.showHelpOnFailure]
      },
      {
        option: '-v, --value <value>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.key !== settingsNames.showHelpOnFailure) {
      return `${args.options.key} is not a valid value for the service option. Allowed values: ${settingsNames.showHelpOnFailure}`;
    }

    return true;
  }
}

module.exports = new CliConfigSetCommand();