import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoPageListCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_LIST;
  }

  public get description(): string {
    return 'Lists all modern pages in the given site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'Title'];
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving client-side pages...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/sitepages/pages?$orderby=Title`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    let pages: any[] = [];

    request
      .get<{ value: any[] }>(requestOptions)
      .then((res: { value: any[] }): Promise<any> => {
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        if (res.value && res.value.length > 0) {
          pages = res.value;
        }

        return request.get(requestOptions);
      })
      .then((res: { value: any[] }): void => {
        if (res.value && res.value.length > 0) {
          const clientSidePages: any[] = res.value.filter(p => p.ListItemAllFields.ClientSideApplicationId === 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec');
          pages = pages.map(p => {
            const clientSidePage = clientSidePages.find(cp => cp && cp.ListItemAllFields && cp.ListItemAllFields.Id === p.Id);
            if (clientSidePage) {
              return {
                ...clientSidePage,
                ...p
              };
            }

            return p;
          });

          logger.log(pages);
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoPageListCommand();