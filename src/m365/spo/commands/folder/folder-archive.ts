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

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { webUrl, url, id, force } = args.options;

    if (!force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you would like to archive this item? You will be able to reactivate it instantly for the first 7 days. After that, it will take up to 24 hours to reactivate.` });
      if (!result) {
        return;
      }
    }

    try {
      if (this.verbose) {
        await logger.logToStderr(`Archiving folder ${url || id} at site ${webUrl}...`);
      }

      let requestUrl: string = `${webUrl}/_api/web`;

      if (id) {
        requestUrl += `/GetFolderById('${formatting.encodeQueryParameter(id)}')`;
      }
      else if (url) {
        const serverRelativePath = urlUtil.getServerRelativePath(webUrl, url);
        requestUrl += `/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
      }
      requestUrl += '?$select=Exists,ListItemAllFields/Id,ListItemAllFields/ParentList/Id&$expand=ListItemAllFields,ListItemAllFields/ParentList';

      const folderInfo = await request.get<{ Exists?: boolean; ListItemAllFields?: { Id: number; ParentList: { Id: string } } }>({
        url: requestUrl,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      });

      if (!folderInfo.Exists) {
        throw `The folder '${url || id}' does not exist.`;
      }

      if (!folderInfo.ListItemAllFields?.ParentList) {
        throw `The folder '${url || id}' is the root folder of a document library and cannot be archived. Archive a subfolder instead.`;
      }

      const requestOptions: CliRequestOptions = {
        url: `${webUrl}/_api/Lists(guid'${folderInfo.ListItemAllFields.ParentList.Id}')/items(${folderInfo.ListItemAllFields.Id})/Archive`,
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