import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { PageTemplateResponse } from './PageTemplateResponse';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving templates...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/sitepages/pages/templates`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<PageTemplateResponse>(requestOptions)
      .then((res: PageTemplateResponse): void => {
        if (res.value && res.value.length > 0) {
          logger.log(res.value);
        }

        cb();
      })
      .catch((err: any): void => {
        // The API returns a 404 when no templates are created on the site collection
        if (err && err.response && err.response.status && err.response.status === 404) {
          logger.log([]);
          cb();
          return;
        }

        return this.handleRejectedODataJsonPromise(err, logger, cb);
      });
  }
}

module.exports = new SpoPageTemplateListCommand();