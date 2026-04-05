import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { validation } from '../../../../utils/validation.js';
import { cli } from '../../../../cli/cli.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  url: z.string().optional(),
  id: z.uuid().optional().alias('i'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileArchiveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_ARCHIVE;
  }

  public get description(): string {
    return 'Archives a file';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.url, options.id].filter(o => o !== undefined).length === 1, {
        error: `Specify 'url' or 'id', but not both.`
      });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { webUrl, url, id, force, verbose } = args.options;

    if (!force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to archive the file ${url || id} at site ${webUrl}?` });
      if (!result) {
        return;
      }
    }

    try {
      if (verbose) {
        await logger.logToStderr(`Archiving file ${url || id} at site ${webUrl}...`);
      }

      let requestUrl: string = '';

      if (id) {
        requestUrl = `${webUrl}/_api/web/GetFileById('${formatting.encodeQueryParameter(id)}')`;
      }
      else if (url) {
        requestUrl = `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl=@f)`;
      }

      let options: string = '?$select=ListId&$expand=ListItemAllFields';

      if (url) {
        const serverRelativePath = urlUtil.getServerRelativePath(webUrl, url);
        options += `&@f='${formatting.encodeQueryParameter(serverRelativePath)}'`;
      }

      const fileInfo = await request.get<{ ListId: string; ListItemAllFields: { Id: number } }>({
        url: requestUrl + options,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      });

      const archiveUrl = `${webUrl}/_api/Lists(guid'${fileInfo.ListId}')/items(${fileInfo.ListItemAllFields.Id})/Archive`;
      const requestOptions: CliRequestOptions = {
        url: archiveUrl,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFileArchiveCommand();