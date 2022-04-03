import { Cli, CommandOutput, Logger } from '../../../cli';
import Command from '../../../Command';
import * as AadAppGetCommand from '../../aad/commands/app/app-get';
import { Options as AadAppGetCommandOptions } from '../../aad/commands/app/app-get';
import AppCommand, { AppCommandArgs } from '../../base/AppCommand';
import commands from '../commands';

class AppGetCommand extends AppCommand {
  public get name(): string {
    return commands.GET;
  }

  public get description(): string {
    return 'Retrieves information about the current Azure AD app';
  }

  public commandAction(logger: Logger, args: AppCommandArgs, cb: (err?: any) => void): void {
    const options: AadAppGetCommandOptions = {
      appId: this.appId,
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose
    };

    Cli
      .executeCommandWithOutput(AadAppGetCommand as Command, { options: { ...options, _: [] } })
      .then((appGetOutput: CommandOutput): void => {
        if (this.verbose) {
          logger.logToStderr(appGetOutput.stderr);
        }

        logger.log(JSON.parse(appGetOutput.stdout));
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AppGetCommand();