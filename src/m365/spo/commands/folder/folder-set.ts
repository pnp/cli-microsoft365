import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FolderColorValues } from './FolderColor.js';

const folderColorValues = [...Object.keys(FolderColorValues), ...Object.values(FolderColorValues)] as [string, ...string[]];

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  url: z.string(),
  name: z.string().optional().alias('n'),
  color: z.enum(folderColorValues, {
    error: e => `'${e.input}' is not a valid value for option 'color'. Allowed values are ${Object.keys(FolderColorValues).join(', ')}, ${Object.values(FolderColorValues).join(', ')}.`
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderSetCommand extends SpoCommand {

  public get name(): string {
    return commands.FOLDER_SET;
  }

  public get description(): string {
    return 'Updates a folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.name !== undefined || opts.color !== undefined, {
        error: `Specify at least one of the options: name or color.`,
        params: {
          customCode: 'required'
        }
      });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Updating folder '${args.options.name}'...`);
      }

      const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      if (!args.options.color) {
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/Web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/ListItemAllFields`,
          headers: {
            accept: 'application/json;odata=nometadata',
            'if-match': '*'
          },
          data: {
            FileLeafRef: args.options.name,
            Title: args.options.name
          },
          responseType: 'json'
        };

        const response = await request.patch<any>(requestOptions);
        if (response && response['odata.null'] === true) {
          throw 'Folder not found.';
        }
      }
      else {
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/foldercoloring/${args.options.name ? 'renamefolder' : 'stampcolor'}(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            coloringInformation: {
              ColorHex: FolderColorValues[args.options.color!] || args.options.color
            },
            newName: args.options.name
          }
        };
        await request.post(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderSetCommand();