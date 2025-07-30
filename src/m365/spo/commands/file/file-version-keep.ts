import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { validation } from '../../../../utils/validation.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = globalOptionsZod
  .extend({
    webUrl: zod.alias('u', z.string()
      .refine(url => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      }))
    ),
    fileUrl: z.string().optional(),
    fileId: zod.alias('i', z.string().optional()
      .refine(id => id === undefined || validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      }))
    ),
    label: z.string()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileVersionKeepCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_VERSION_KEEP;
  }

  public get description(): string {
    return 'Ensure that a specific file version will never expire';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => (options.fileUrl !== undefined) !== (options.fileId !== undefined), {
        message: `Specify 'fileUrl' or 'fileId', but not both.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Ensuring version '${args.options.label}' of file '${args.options.fileUrl || args.options.fileId}' at site '${args.options.webUrl}' will never expire...`);
    }

    try {
      const fileUrl = this.getFileUrl(args.options.webUrl, args.options.fileUrl, args.options.fileId);

      const requestVersionOptions: CliRequestOptions = {
        url: `${fileUrl}/versions/?$filter=VersionLabel eq '${args.options.label}'`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const response = await request.get<{ value: { ID: number }[] }>(requestVersionOptions);
      const version: { ID: number; } | undefined = response.value[0];

      if (version === undefined) {
        throw `Version with label '${args.options.label}' not found.`;
      }

      const requestExpirationOptions: CliRequestOptions = {
        url: `${fileUrl}/versions(${version.ID})/SetExpirationDate()`,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      await request.post(requestExpirationOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getFileUrl(webUrl: string, fileUrl?: string, fileId?: string): string {
    let requestUrl: string;

    if (fileUrl) {
      const serverRelUrl = urlUtil.getServerRelativePath(webUrl, fileUrl);
      requestUrl = `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelUrl)}')`;
    }
    else {
      requestUrl = `${webUrl}/_api/web/GetFileById('${fileId}')`;
    }

    return requestUrl;
  }
}

export default new SpoFileVersionKeepCommand();