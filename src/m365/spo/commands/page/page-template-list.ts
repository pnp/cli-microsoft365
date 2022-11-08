import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { PageTemplate } from './PageTemplate';

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
      logger.logToStderr(`Retrieving templates...`);
    }

    try {
      const res = await odata.getAllItems<PageTemplate>(`${args.options.webUrl}/_api/sitepages/pages/templates`);
      if (res && res.length > 0) {
        logger.log(res);
      }
    }
    catch (err: any) {
      // The API returns a 404 when no templates are created on the site collection
      if (err && err.response && err.response.status && err.response.status === 404) {
        logger.log([]);
        return;
      }

      return this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoPageTemplateListCommand();