import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  label: z.string(),
  fileUrl: z.string().optional(),
  fileId: z.uuid().optional().alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileVersionGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_VERSION_GET;
  }

  public get description(): string {
    return 'Gets information about a specific version of a specified file';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.fileUrl, opts.fileId].filter(o => o !== undefined).length === 1, {
        error: `Specify 'fileUrl' or 'fileId', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['fileUrl', 'fileId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving version ${args.options.label} of the file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
    }

    try {
      const version = await this.getVersion(args);
      await logger.log(version.value[0]);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getVersion(args: CommandArgs): Promise<any> {
    let requestUrl: string = `${args.options.webUrl}/_api/web/`;
    if (args.options.fileUrl) {
      const serverRelUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
      requestUrl += `GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelUrl)}')/versions/?$filter=VersionLabel eq '${args.options.label}'`;
    }
    else {
      requestUrl += `GetFileById('${args.options.fileId}')/versions/?$filter=VersionLabel eq '${args.options.label}'`;
    }

    requestUrl += `&$select=*,ExpirationDate`;

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: any[] }>(requestOptions);
    return response;
  }
}

export default new SpoFileVersionGetCommand();