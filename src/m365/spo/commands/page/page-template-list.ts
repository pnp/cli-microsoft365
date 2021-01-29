import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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

class SpoPageControlListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_TEMPLATE_LIST}`;
  }

  public get description(): string {
    return 'Lists all page templates in the given site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'FileName', 'Id', 'PageLayoutType', 'Url'];
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

        return this.handleRejectedODataJsonPromise(err, logger, cb)
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageControlListCommand();