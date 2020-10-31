import auth from '../../../Auth';
import { Logger } from '../../../cli';
import GlobalOptions from '../../../GlobalOptions';
import SpoCommand from '../../base/SpoCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions { }

interface SpoContext {
  SpoUrl: string;
}

class SpoGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.GET}`;
  }

  public get description(): string {
    return 'Gets the context URL for the root SharePoint site collection and SharePoint tenant admin site';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const spoContext: SpoContext = {
      SpoUrl: auth.service.spoUrl ? auth.service.spoUrl : ''
    };
    logger.log(spoContext);
    cb();
  }
}

module.exports = new SpoGetCommand();