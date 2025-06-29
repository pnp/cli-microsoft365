import { z } from 'zod';
import { Logger } from '../../../cli/Logger.js';
import { entraApp } from '../../../utils/entraApp.js';
import AppCommand, { AppCommandArgs, appCommandOptions } from '../../base/AppCommand.js';
import commands from '../commands.js';

class AppGetCommand extends AppCommand {
  public get name(): string {
    return commands.GET;
  }

  public get description(): string {
    return 'Retrieves information about the current Microsoft Entra app';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return appCommandOptions;
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