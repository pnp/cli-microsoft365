import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  ids: z.string()
    .refine(ids => formatting.splitAndTrim(ids).every(id => validation.isValidGuid(id)), {
      error: e => `The following GUIDs are invalid for the option 'ids': ${validation.isValidGuidArray(e.input as string)}.`
    })
    .optional()
    .alias('i'),
  allPrimary: z.boolean().optional(),
  allSecondary: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteRecycleBinItemRestoreCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores given items from the site recycle bin';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => o.ids !== undefined || o.allPrimary || o.allSecondary, {
        error: `Specify either 'ids', 'allPrimary', or 'allSecondary'.`,
        params: {
          customCode: 'optionSet',
          options: ['ids', 'allPrimary', 'allSecondary']
        }
      })
      .refine(o => !(o.ids !== undefined && (o.allPrimary || o.allSecondary)), {
        error: `Option 'ids' cannot be used with 'allPrimary' or 'allSecondary'.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Restoring items from recycle bin at ${args.options.siteUrl}...`);
    }

    const baseUrl: string = `${args.options.siteUrl}/_api`;

    try {
      if (args.options.ids) {
        const requestUrl = baseUrl + '/site/RecycleBin/RestoreByIds';
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
          await this.restoreRecycleBinStage(baseUrl + '/site/GetRecycleBinItems(rowLimit=2000000000,itemState=2)/RestoreAll');
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
