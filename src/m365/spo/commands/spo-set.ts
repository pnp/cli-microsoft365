import auth from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import GlobalOptions from '../../../GlobalOptions.js';
import { urlUtil } from '../../../utils/urlUtil.js';
import { validation } from '../../../utils/validation.js';
import SpoCommand from '../../base/SpoCommand.js';
import commands from '../commands.js';

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
    this.#initTypes();
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

  #initTypes(): void {
    this.types.string.push('url');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    auth.connection.spoUrl = urlUtil.removeTrailingSlashes(args.options.url);

    try {
      await auth.storeConnectionInfo();
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSetCommand();