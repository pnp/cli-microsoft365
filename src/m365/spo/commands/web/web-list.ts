import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { WebProperties } from './WebProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
}

class SpoWebListCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_LIST;
  }

  public get description(): string {
    return 'Lists subsites of the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url', 'Id'];
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving all webs in site at ${args.options.url}...`);
    }

    let requestUrl: string = `${args.options.url}/_api/web/webs`;

    if (args.options.output !== 'json') {
      requestUrl += '?$select=Title,Id,URL';
    }

    try {
      const webProperties: WebProperties[] = await odata.getAllItems<WebProperties>(requestUrl);
      await logger.log(webProperties);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebListCommand();