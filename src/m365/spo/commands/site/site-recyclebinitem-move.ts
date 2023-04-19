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
  all?: boolean;
  confirm?: boolean;
}

class SpoSiteRecycleBinItemMoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_MOVE;
  }

  public get description(): string {
    return 'Moves items from the first-stage recycle bin to the second-stage recycle bin';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        ids: typeof args.options.ids !== 'undefined',
        all: !!args.options.all,
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
        option: '--all'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['ids', 'all'] }
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
      logger.logToStderr(`Moving items from the first-stage recycle bin to the second-stage recycle bin at ${args.options.siteUrl}...`);
    }

    if (args.options.confirm) {
      await this.moveRecycleBinItem(args, logger);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: 'Are you sure you want to move the items to the second-stage recycle bin?'
      });

      if (result.continue) {
        await this.moveRecycleBinItem(args, logger);
      }
    }
  }

  private async moveRecycleBinItem(args: CommandArgs, logger: Logger): Promise<void> {
    try {
      if (args.options.all) {
        if (this.verbose) {
          logger.logToStderr('Moving all items to the second-stage recycle bin');
        }
        const requestOptions: CliRequestOptions = {
          url: `${args.options.siteUrl}/_api/web/recycleBin/MoveAllToSecondStage`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
        const response = await request.post<{ 'odata.null': boolean }>(requestOptions);
        if (!response['odata.null']) {
          throw 'Something went wrong when moving the selected items to the second-stage recycle bin';
        }
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`Moving ${args.options.ids} to the second-stage recycle bin`);
        }

        await this.postBatch(args.options.ids!.split(','), logger, args.options.siteUrl);
      }
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
      batchContents.push(`POST ${siteUrl.replace(/\/$/, '')}/_api/web/recycleBin('${id.trim()}')/MoveToSecondStage HTTP/1.1`);
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
      throw `Something went wrong while moving the selected item(s) to the second-stage recycle bin: ${errors.join(', ')}`;
    }
  }
}

module.exports = new SpoSiteRecycleBinItemMoveCommand();