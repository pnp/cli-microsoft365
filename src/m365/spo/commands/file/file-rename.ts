import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import removeCommand, { Options as SpoFileRemoveOptions } from './file-remove.js';
import { FileProperties } from './FileProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  sourceUrl: z.string().alias('s'),
  targetFileName: z.string().alias('t'),
  force: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface RenameResponse {
  value: RenameResponseValue[];
}

interface RenameResponseValue {
  ErrorCode: number;
  ErrorMessage: string;
  FieldName: string;
  FieldValue: string;
  HasException: boolean;
  ItemId: number;
}

class SpoFileRenameCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_RENAME;
  }

  public get description(): string {
    return 'Renames a file';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const webUrl = args.options.webUrl;
    const originalFileServerRelativePath: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.sourceUrl);

    try {
      await this.getFile(originalFileServerRelativePath, webUrl);

      if (args.options.force) {
        await this.deleteFile(webUrl, args.options.sourceUrl, args.options.targetFileName);
      }

      const requestBody: any = {
        formValues: [{
          FieldName: 'FileLeafRef',
          FieldValue: args.options.targetFileName
        }]
      };

      const requestOptions: CliRequestOptions = {
        url: `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(originalFileServerRelativePath)}')/ListItemAllFields/ValidateUpdateListItem()`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      const resp = await request.post<RenameResponse>(requestOptions);
      await logger.log(resp.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFile(originalFileServerRelativeUrl: string, webUrl: string): Promise<FileProperties> {
    const requestUrl = `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(originalFileServerRelativeUrl)}')?$select=UniqueId`;
    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    return request.get(requestOptions);
  }

  private async deleteFile(webUrl: string, sourceUrl: string, targetFileName: string): Promise<void> {
    const targetFileServerRelativeUrl: string = `${urlUtil.getServerRelativePath(webUrl, sourceUrl.substring(0, sourceUrl.lastIndexOf('/')))}/${targetFileName}`;

    const removeOptions: SpoFileRemoveOptions = {
      webUrl: webUrl,
      url: targetFileServerRelativeUrl,
      force: true,
      debug: this.debug,
      verbose: this.verbose
    };
    try {
      await cli.executeCommand(removeCommand as Command, { options: { ...removeOptions, _: [] } });
    }
    catch (err: any) {
      if (err?.error?.message?.includes('does not exist')) {
        return;
      }
      throw err;
    }
  }
}

export default new SpoFileRenameCommand();
