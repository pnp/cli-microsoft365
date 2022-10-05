import { Cli } from '../../../cli/Cli';
import { Logger } from '../../../cli/Logger';
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

  public async commandAction(logger: Logger, args: AppCommandArgs): Promise<void> {
    const options: AadAppGetCommandOptions = {
      appId: this.appId,
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose
    };

    try {
      const appGetOutput = await Cli.executeCommandWithOutput(AadAppGetCommand as Command, { options: { ...options, _: [] } });
      if (this.verbose) {
        logger.logToStderr(appGetOutput.stderr);
      }

      logger.log(JSON.parse(appGetOutput.stdout));
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AppGetCommand();