import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import Command, {
  CommandOption,
  CommandError
} from '../../../../Command';
import auth from '../../../../Auth';
import { CommandInstance } from '../../../../cli';

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
}

module.exports = new AccessTokenGetCommand();