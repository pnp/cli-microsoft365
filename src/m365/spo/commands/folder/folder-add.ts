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
import { FolderProperties } from './FolderProperties.js';

const folderColorValues = [...Object.keys(FolderColorValues), ...Object.values(FolderColorValues)] as [string, ...string[]];

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  parentFolderUrl: z.string().alias('p'),
  name: z.string().alias('n'),
  color: z.enum(folderColorValues, {
    error: e => `'${e.input}' is not a valid value for option 'color'. Allowed values are ${Object.keys(FolderColorValues).join(', ')}, ${Object.values(FolderColorValues).join(', ')}.`
  }).optional(),
  ensureParentFolders: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderAddCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ADD;
  }

  public get description(): string {
    return 'Creates a folder within a parent folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const parentFolderServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.parentFolderUrl);
      const serverRelativeUrl: string = `${parentFolderServerRelativeUrl}/${args.options.name}`;

      if (args.options.ensureParentFolders) {
        await this.ensureParentFolderPath(args.options, parentFolderServerRelativeUrl, logger);
      }

      const folder = await this.addFolder(serverRelativeUrl, args.options, logger);
      await logger.log(folder);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async ensureParentFolderPath(options: Options, parentFolderPath: string, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Ensuring parent folders exist...`);
    }

    const parsedUrl = new URL(options.webUrl);
    const absoluteFolderUrl: string = `${parsedUrl.origin}${parentFolderPath}`;
    const relativeFolderPath = absoluteFolderUrl.replace(options.webUrl, '');

    const parentFolders: string[] = relativeFolderPath.split('/').filter(folder => folder !== '');

    for (let i = 1; i < parentFolders.length; i++) {
      const currentFolderPath = parentFolders.slice(0, i + 1).join('/');

      if (this.verbose) {
        await logger.logToStderr(`Checking if folder '${currentFolderPath}' exists...`);
      }

      const folderExists = await this.getFolderExists(options.webUrl, currentFolderPath);
      if (!folderExists) {
        await this.addFolder(currentFolderPath, options, logger);
      }
    }
  }

  private async getFolderExists(webUrl: string, folderServerRelativeUrl: string): Promise<boolean> {
    const requestUrl = `${webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderServerRelativeUrl)}')?$select=Exists`;

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ Exists: boolean }>(requestOptions);
    return response.Exists;
  }

  private async addFolder(serverRelativeUrl: string, options: Options, logger: Logger): Promise<FolderProperties> {
    if (this.verbose) {
      const folderName = serverRelativeUrl.split('/').pop();
      const folderLocation = serverRelativeUrl.split('/').slice(0, -1).join('/');
      await logger.logToStderr(`Adding folder with name '${folderName}' at location '${folderLocation}'...`);
    }

    const requestOptions: CliRequestOptions = {
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    if (options.color === undefined) {
      requestOptions.url = `${options.webUrl}/_api/web/folders/addUsingPath(decodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    }
    else {
      requestOptions.url = `${options.webUrl}/_api/foldercoloring/createfolder(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}', overwrite=false)`;
      requestOptions.data = {
        coloringInformation: {
          ColorHex: FolderColorValues[options.color] || options.color
        }
      };
    }

    const response = await request.post<FolderProperties>(requestOptions);
    return response;
  }
}

export default new SpoFolderAddCommand();