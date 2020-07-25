import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../Command';
import SpoCommand from '../../base/SpoCommand';
import auth from '../../../Auth';
import { CommandInstance } from '../../../cli';

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
      return SpoCommand.isValidSharePointUrl(args.options.url);
    };
  }
}

module.exports = new SpoSetCommand();