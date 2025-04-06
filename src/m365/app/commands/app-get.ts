import { Logger } from '../../../cli/Logger.js';
import AppCommand, { AppCommandArgs } from '../../base/AppCommand.js';
import commands from '../commands.js';
import { entraApp } from '../../../utils/entraApp.js';

class AppGetCommand extends AppCommand {
  public get name(): string {
    return commands.GET;
  }

  public get description(): string {
    return 'Retrieves information about the current Microsoft Entra app';
  }

  public async commandAction(logger: Logger, args: AppCommandArgs): Promise<void> {
    try {
      const app = await entraApp.getAppRegistrationByAppId(args.options.appId!);
      await logger.log(app);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AppGetCommand();