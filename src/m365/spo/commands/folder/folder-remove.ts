import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  url: string;
  recycle?: boolean;
  force?: boolean;
}

class SpoFolderRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_REMOVE;
  }

  public get description(): string {
    return 'Deletes the specified folder';
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
        recycle: !!args.options.recycle,
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
        option: '--url <url>'
      },
      {
        option: '--recycle'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeFolder(logger, args.options);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the folder ${args.options.url} located in site ${args.options.webUrl}?`
      });

      if (result.continue) {
        await this.removeFolder(logger, args.options);
      }
    }
  }

  private async removeFolder(logger: Logger, options: Options): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing folder in site at ${options.webUrl}...`);
    }

    const serverRelativePath: string = urlUtil.getServerRelativePath(options.webUrl, options.url);
    let requestUrl: string = `${options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
    if (options.recycle) {
      requestUrl += `/recycle()`;
    }
    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'If-Match': '*',
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoFolderRemoveCommand();