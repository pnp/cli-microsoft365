import auth from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import SpoCommand from '../../base/SpoCommand.js';
import commands from '../commands.js';

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

  public async commandAction(logger: Logger): Promise<void> {
    const spoContext: SpoContext = {
      SpoUrl: auth.service.spoUrl ? auth.service.spoUrl : ''
    };
    await logger.log(spoContext);
  }
}

export default new SpoGetCommand();