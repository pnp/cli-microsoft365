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
    this.#initOptionSets();
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

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['ids'],
        runsWhen: (args) => args.options.allPrimary !== undefined && args.options.allSecondary !== undefined
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Restoring items from recycle bin at ${args.options.siteUrl}...`);
    }

    let requestUrl: string = `${args.options.siteUrl}/_api`;

    if (args.options.ids) {
      requestUrl += `/site/RecycleBin/RestoreByIds`;
    }
    else if (args.options.allPrimary) {
      requestUrl += `/web/RecycleBin/RestoreAll`;
    }
    else if (args.options.allSecondary) {
      requestUrl += `/site/RecycleBin/RestoreAll`;
    }

    try {
      if (args.options.ids) {
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
        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata',
            'content-type': 'application/json'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteRecycleBinItemRestoreCommand();