import auth from '../../../Auth.js';
import { globalOptionsZod } from '../../../Command.js';
import { Logger } from '../../../cli/Logger.js';
import SpoCommand from '../../base/SpoCommand.js';
import commands from '../commands.js';
import { z } from 'zod';

export const options = globalOptionsZod.strict();

interface SpoContext {
  SpoUrl: string;
}

class SpoGetCommand extends SpoCommand {
  public get name(): string {
    return commands.GET;
  }

  public get description(): string {
    return 'Gets the context URL for the root SharePoint site collection and SharePoint tenant admin site';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger): Promise<void> {
    const spoContext: SpoContext = {
      SpoUrl: auth.connection.spoUrl ? auth.connection.spoUrl : ''
    };
    await logger.log(spoContext);
  }
}

export default new SpoGetCommand();