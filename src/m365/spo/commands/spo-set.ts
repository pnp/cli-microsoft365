import auth from '../../../Auth';
import { Logger } from '../../../cli';
import {
  CommandError
} from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import { validation } from '../../../utils';
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

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    auth.service.spoUrl = args.options.url;
    auth.storeConnectionInfo().then(() => {
      cb();
    }, err => {
      cb(new CommandError(err));
    });
  }
}

module.exports = new SpoSetCommand();