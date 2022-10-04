import auth from '../../../Auth';
import { Logger } from '../../../cli/Logger';
import SpoCommand from '../../base/SpoCommand';
import commands from '../commands';

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
    logger.log(spoContext);    
  }
}

module.exports = new SpoGetCommand();