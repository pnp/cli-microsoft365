import auth from '../../../Auth';
import { Logger } from '../../../cli';
import {
  CommandError, CommandOption
} from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import SpoCommand from '../../base/SpoCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
}

class SpoSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SET;
  }

  public get description(): string {
    return 'Sets the URL of the root SharePoint site collection for use in SPO commands';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
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
        option: '-u, --url <url>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.url);
  }
}

module.exports = new SpoSetCommand();