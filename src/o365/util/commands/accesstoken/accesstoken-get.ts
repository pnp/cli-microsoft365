import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import Command, {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import auth from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  new?: boolean;
  resource: string;
}

class AccessTokenGetCommand extends Command {
  public get name(): string {
    return `${commands.UTIL_ACCESSTOKEN_GET}`;
  }

  public get description(): string {
    return 'Gets access token for the specified resource';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .ensureAccessToken(args.options.resource, cmd, this.debug, args.options.new)
      .then((accessToken: string): void => {
        cmd.log(accessToken);
        cb();
      }, (err: any): void => cb(new CommandError(err)));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-r, --resource <resource>',
        description: 'The resource for which to retrieve an access token'
      },
      {
        option: '--new',
        description: 'Retrieve a new access token to ensure that it\'s valid for as long as possible'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.resource) {
        return 'Required parameter resource missing';
      }

      return true;
    };
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
    
    The ${chalk.blue(this.name)} command returns an access token for the specified
    resource. If an access token has been previously retrieved and is still
    valid, the command will return the cached token. If you want to ensure that
    the returned access token is valid for as long as possible, you can force
    the command to retrieve a new access token by using the ${chalk.grey('--new')} option.
      
  Examples:
  
    Get access token for the Microsoft Graph
      ${this.name} --resource https://graph.microsoft.com

    Get a new access token for SharePoint Online
      ${this.name} --resource https://contoso.sharepoint.com --new
`);
  }
}

module.exports = new AccessTokenGetCommand();