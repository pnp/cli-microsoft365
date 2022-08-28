import auth from '../../../Auth';
import { Logger } from '../../../cli';
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
    try {
      const spoContext: SpoContext = {
        SpoUrl: auth.service.spoUrl ? auth.service.spoUrl : ''
      };
      logger.log(spoContext);      
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoGetCommand();