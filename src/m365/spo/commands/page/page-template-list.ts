import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { PageTemplate } from './PageTemplate.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
}

class SpoPageTemplateListCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_TEMPLATE_LIST;
  }

  public get description(): string {
    return 'Lists all page templates in the given site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'FileName', 'Id', 'PageLayoutType', 'Url'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving templates...`);
    }

    try {
      const res = await odata.getAllItems<PageTemplate>(`${args.options.webUrl}/_api/sitepages/pages/templates`);
      if (res && res.length > 0) {
        await logger.log(res);
      }
    }
    catch (err: any) {
      // The API returns a 404 when no templates are created on the site collection
      if (err && err.response && err.response.status && err.response.status === 404) {
        await logger.log([]);
        return;
      }

      return this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageTemplateListCommand();