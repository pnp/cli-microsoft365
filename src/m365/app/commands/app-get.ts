import { Logger } from '../../../cli/Logger.js';
import { aadApp } from '../../../utils/aadApp.js';
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
    try {
      const app = await aadApp.getAppById(args.options.appId!);
      logger.log(app);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AppGetCommand();