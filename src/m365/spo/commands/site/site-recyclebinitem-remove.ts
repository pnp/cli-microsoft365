import { v4 } from 'uuid';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  ids?: string;
  confirm?: boolean;
}

class SpoSiteRecycleBinItemRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Permanently deletes specific items from the site recycle bin';
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
        ids: typeof args.options.ids !== 'undefined',
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --ids [ids]'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.ids && !validation.isValidGuidArray(args.options.ids.split(','))) {
          return 'ids contains invalid GUID';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Permanently deleting specific items from the site recycle bin at ${args.options.siteUrl}...`);
    }

    if (args.options.confirm) {
      await this.removeRecycleBinItem(args, logger);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: 'Are you sure you want to permanently delete the items from the site recycle bin?'
      });

      if (result.continue) {
        await this.removeRecycleBinItem(args, logger);
      }
    }
  }

  private async removeRecycleBinItem(args: CommandArgs, logger: Logger): Promise<void> {
    try {
      await this.postBatch(args.options.ids!.split(','), logger, args.options.siteUrl);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async postBatch(ids: string[], logger: Logger, siteUrl: string): Promise<void> {
    const errors: string[] = [];
    const batchGuid = v4();
    const changeSetId = v4();
    const batchContents: string[] = [];

    batchContents.push(`--batch_${batchGuid}`);
    batchContents.push(`Content-Type: multipart/mixed; boundary="changeset_${changeSetId}"`);
    batchContents.push('Content-Transfer-Encoding: binary');
    batchContents.push('');

    ids.forEach((id) => {
      batchContents.push(`--changeset_${changeSetId}`);
      batchContents.push('Content-Type: application/http');
      batchContents.push('Content-Transfer-Encoding: binary');
      batchContents.push('');
      batchContents.push(`POST ${siteUrl.replace(/\/$/, '')}/_api/web/recycleBin('${id.trim()}')/DeleteObject HTTP/1.1`);
      batchContents.push(`Accept: application/json;odata=verbose`);
      batchContents.push('');
    });

    batchContents.push(`--changeset_${changeSetId}--`);
    batchContents.push(`--batch_${batchGuid}--`);

    if (this.verbose) {
      logger.logToStderr(`Batchbody: ${batchContents}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${siteUrl.replace(/\/$/, '')}/_api/$batch`,
      headers: {
        'Content-Type': `multipart/mixed; boundary=batch_${batchGuid}`,
        'Accept': 'application/json;odata=verbose'
      },
      data: batchContents.join('\r\n'),
      responseType: 'json'
    };
    const response: string = await request.post(requestOptions);
    const responseInLines = response.replace(/[\r\n\\]+/g, '\n').split('\n');
    for (let currentLine = 0; currentLine < responseInLines.length; currentLine++) {
      try {
        // parse the JSON response...
        const line = responseInLines[currentLine];
        const tryParseJson = JSON.parse(line);
        if (tryParseJson.error) {
          errors.push(tryParseJson.error.message.value);
        }
      }
      catch (e) {
        // don't do anything... just keep moving
      }
    }

    if (errors.length > 0) {
      throw `Something went wrong while permanently deleting the selected item(s) from the recycle bin: ${errors.join(', ')}`;
    }
  }
}

module.exports = new SpoSiteRecycleBinItemRemoveCommand();