import * as chalk from "chalk";
import { Logger } from "../../../../cli";
import { CommandOption } from "../../../../Command";
import configstore from "../../../../configstore";
import GlobalOptions from "../../../../GlobalOptions";
import AnonymousCommand from "../../../base/AnonymousCommand";
import commands from "../../commands";
const Configstore = require('configstore');

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let value: any = '';
    const conf = new Configstore(configstore.name);

    switch (args.options.key) {
      case configstore.showHelpOnFailure:
        value = args.options.value === "true";
        break;
    }

    conf.set(args.options.key, value);
    
    if (this.verbose) {
      logger.logToStderr(chalk.green('DONE'));
    }

    cb();
  }

  public action(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this.initAction(args, logger);
    this.commandAction(logger, args, cb);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-k, --key <key>',
        autocomplete: [configstore.showHelpOnFailure]
      },
      {
        option: '-v, --value <value>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.key !== configstore.showHelpOnFailure) {
      return `${args.options.key} is not a valid value for the service option. Allowed values: ${configstore.showHelpOnFailure}`;
    }

    return true;
  }
}

module.exports = new CliConfigSetCommand();