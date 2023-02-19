import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import removeCommand, { Options as SpoFileRemoveOptions } from './file-remove.js';
import { FileProperties } from './FileProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetFileName: string;
  force?: boolean;
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-s, --sourceUrl <sourceUrl>'
      },
      {
        option: '-t, --targetFileName <targetFileName>'
      },
      {
        option: '--force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
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
        url: `${webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(originalFileServerRelativePath)}')/ListItemAllFields/ValidateUpdateListItem()`,
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
    const requestUrl = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(originalFileServerRelativeUrl)}')?$select=UniqueId`;
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
      recycle: true,
      force: true,
      debug: this.debug,
      verbose: this.verbose
    };
    try {
      await Cli.executeCommand(removeCommand as Command, { options: { ...removeOptions, _: [] } });
    }
    catch (err: any) {
      if (err.error !== undefined && err.error.message !== undefined && err.error.message.includes('does not exist')) {

      }
      else {
        throw err;
      }
    }
  }
}

export default new SpoFileRenameCommand();
