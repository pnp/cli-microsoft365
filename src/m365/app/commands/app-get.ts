import { Cli } from '../../../cli/Cli.js';
import { Logger } from '../../../cli/Logger.js';
import Command from '../../../Command.js';
import aadAppGetCommand, { Options as AadAppGetCommandOptions } from '../../aad/commands/app/app-get.js';
import AppCommand, { AppCommandArgs } from '../../base/AppCommand.js';
import commands from '../commands.js';

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
      const appGetOutput = await Cli.executeCommandWithOutput(aadAppGetCommand as Command, { options: { ...options, _: [] } });
      if (this.verbose) {
        await logger.logToStderr(appGetOutput.stderr);
      }

      await logger.log(JSON.parse(appGetOutput.stdout));
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AppGetCommand();