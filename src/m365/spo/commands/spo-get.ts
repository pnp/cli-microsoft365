import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../Command';
import SpoCommand from '../../base/SpoCommand';
import auth from '../../../Auth';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {}

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const spoContext: SpoContext = {
      SpoUrl: auth.service.spoUrl ? auth.service.spoUrl : ''
    };
    cmd.log(spoContext);
    cb();
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return true;
    };
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    log(vorpal.find(commands.GET).helpInformation());
    log(` Remarks:

    CLI for Microsoft 365 automatically discovers the URL of the root SharePoint site
    collection/SharePoint tenant admin site (whichever is needed to run
    the particular command). Using this command you can see which URLs the CLI has discovered.
      
  Examples:
  
    Get the context URL for the root SharePoint site collection and SharePoint tenant admin site
      ${commands.GET} --output json
`);
  }
}

module.exports = new SpoGetCommand();