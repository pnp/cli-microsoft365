import * as chalk from "chalk";
import { Logger } from "../../../../cli";
import { CommandOption } from "../../../../Command";
import config from  "../../../../config";
import configstoreOptions from "../../../../configstoreOptions";
import GlobalOptions from "../../../../GlobalOptions";
import AnonymousCommand from "../../../base/AnonymousCommand";
import commands from "../../commands";
import * as Configstore from 'configstore';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  key: string;
  value: string;
}

class CliConfigSetCommand extends AnonymousCommand {

  public get name(): string {
    return `${commands.CONFIG_SET}`;
  }

  public get description(): string {
    return 'Manage global configuration settings about the CLI for Microsoft 365';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.key = args.options.key;
    telemetryProps.value = args.options.value;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let value: any = undefined;
    const conf = new Configstore(config.configstoreName);

    switch (args.options.key) {
      case configstoreOptions.showHelpOnFailure:
        value = args.options.value === "true";
        break;
    }

    conf.set(args.options.key, value);
    
    if (this.verbose) {
      logger.logToStderr(chalk.green('DONE'));
    }

    cb();
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-k, --key <key>',
        autocomplete: [configstoreOptions.showHelpOnFailure]
      },
      {
        option: '-v, --value <value>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.key !== configstoreOptions.showHelpOnFailure) {
      return `${args.options.key} is not a valid value for the service option. Allowed values: ${configstoreOptions.showHelpOnFailure}`;
    }

    return true;
  }
}

module.exports = new CliConfigSetCommand();