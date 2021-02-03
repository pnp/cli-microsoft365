import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, {
    CommandError, CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  new?: boolean;
  resource: string;
}

class AccessTokenGetCommand extends Command {
  public get name(): string {
    return commands.UTIL_ACCESSTOKEN_GET;
  }

  public get description(): string {
    return 'Gets access token for the specified resource';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .ensureAccessToken(args.options.resource, logger, this.debug, args.options.new)
      .then((accessToken: string): void => {
        logger.log(accessToken);
        cb();
      }, (err: any): void => cb(new CommandError(err)));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-r, --resource <resource>'
      },
      {
        option: '--new'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new AccessTokenGetCommand();