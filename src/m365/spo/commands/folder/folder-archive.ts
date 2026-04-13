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

class SpoFolderArchiveCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ARCHIVE;
  }

  public get description(): string {
    return 'Archives a folder';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { webUrl, url, id, force, verbose } = args.options;

    if (!force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to archive the folder ${url || id} at site ${webUrl}?` });
      if (!result) {
        return;
      }
    }

    try {
      if (verbose) {
        await logger.logToStderr(`Archiving folder ${url || id} at site ${webUrl}...`);
      }

      let requestUrl: string = `${webUrl}/_api/web`;
      if (id) {
        requestUrl += `/GetFolderById('${formatting.encodeQueryParameter(id)}')`;
      }
      else if (url) {
        const serverRelativePath = urlUtil.getServerRelativePath(webUrl, url);
        requestUrl += `/GetFolderByServerRelativePath(DecodedUrl='${serverRelativePath}')`;
      }

      const queryString: string = '?$select=ListItemAllFields/Id,ListItemAllFields/ParentList/Id&$expand=ListItemAllFields,ListItemAllFields/ParentList';

      const folderInfo = await request.get<{ ListItemAllFields: { Id: number; ParentList: { Id: string } } }>({
        url: requestUrl + queryString,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      });

      const archiveUrl = `${webUrl}/_api/Lists(guid'${folderInfo.ListItemAllFields.ParentList.Id}')/items(${folderInfo.ListItemAllFields.Id})/Archive`;
      const requestOptions: CliRequestOptions = {
        url: archiveUrl,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderArchiveCommand();