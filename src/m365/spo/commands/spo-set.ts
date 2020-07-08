import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../Command';
import SpoCommand from '../../base/SpoCommand';
import auth from '../../../Auth';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
}

class SpoSetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SET}`;
  }

  public get description(): string {
    return 'Sets the URL of the root SharePoint site collection for use in SPO commands';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    auth.service.spoUrl = args.options.url;
    auth.storeConnectionInfo().then(() => {
      cb();
    }, err => {
      cb(new CommandError(err));
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'The URL of the root SharePoint site collection to use in SPO commands'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.url) {
        return 'Required parameter url missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      return true;
    };
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    log(vorpal.find(commands.SET).helpInformation());
    log(` Remarks:

    CLI for Microsoft 365 automatically discovers the URL of the root SharePoint site
    collection/SharePoint tenant admin site (whichever is needed to run
    the particular command). In specific cases, like when managing multi-geo
    Microsoft 365 tenants, it could be desirable to make the CLI manage
    the specific geography. For such cases, you can use this command
    to explicitly specify the SPO URL that should be used when executing SPO
    commands.
      
  Examples:
  
    Set SPO URL to the specified URL
      ${commands.SET} --url https://contoso.sharepoint.com
`);
  }
}

module.exports = new SpoSetCommand();