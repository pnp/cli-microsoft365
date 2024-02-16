import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  ids?: string;
  allPrimary?: boolean;
  allSecondary?: boolean;
}

class SpoSiteRecycleBinItemRestoreCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores given items from the site recycle bin';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --ids [ids]'
      },
      {
        option: '--allPrimary'
      },
      {
        option: '--allSecondary'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.ids) {
          const invalidIds = formatting
            .splitAndTrim(args.options.ids)
            .filter(id => !validation.isValidGuid(id));

          if (invalidIds.length > 0) {
            return `The following IDs are invalid: ${invalidIds.join(', ')}`;
          }
        }

        if ((!args.options.ids && !args.options.allPrimary && !args.options.allSecondary)
          || (args.options.ids && (args.options.allPrimary || args.options.allSecondary))) {
          return 'Specify ids or allPrimary and/or allSecondary';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Restoring items from recycle bin at ${args.options.siteUrl}...`);
    }

    let baseUrl: string = `${args.options.siteUrl}/_api`;

    try {
      if (args.options.ids) {
        const requestUrl = baseUrl + `/site/RecycleBin/RestoreByIds`;
        const ids: string[] = formatting.splitAndTrim(args.options.ids);
        const idsChunks: string[][] = [];

        while (ids.length) {
          idsChunks.push(ids.splice(0, 20));
        }

        await Promise.all(
          idsChunks.map((idsChunk: string[]) => {
            const requestOptions: CliRequestOptions = {
              url: requestUrl,
              headers: {
                'accept': 'application/json;odata=nometadata',
                'content-type': 'application/json'
              },
              responseType: 'json',
              data: {
                ids: idsChunk
              }
            };

            return request.post(requestOptions);
          })
        );
      }
      else {
        if (args.options.allPrimary && args.options.allSecondary) {
          await this.restoreRecycleBinStage(baseUrl + '/site/RecycleBin/RestoreAll');
        }
        else if (args.options.allPrimary) {
          await this.restoreRecycleBinStage(baseUrl + '/web/RecycleBin/RestoreAll');
        }
        else if (args.options.allSecondary) {
          await this.restoreRecycleBinStage(baseUrl + '/site/GetRecyclebinItems(rowLimit=2000000000,itemState=2)/RestoreAll');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async restoreRecycleBinStage(requestUrl: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }
}

export default new SpoSiteRecycleBinItemRestoreCommand();
