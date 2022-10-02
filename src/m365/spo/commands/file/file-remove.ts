import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  url?: string;
  recycle?: boolean;
  confirm?: boolean;
}

class SpoFileRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified file';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: (!(!args.options.id)).toString(),
        url: (!(!args.options.url)).toString(),
        recycle: (!(!args.options.recycle)).toString(),
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-w, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-u, --url [url]'
      },
      {
        option: '--recycle'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
    
        if (args.options.id &&
          !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }
    
        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'url']);
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeFile: () => Promise<void> = async (): Promise<void> => {
      if (this.verbose) {
        logger.logToStderr(`Removing file in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string = '';

      if (args.options.id) {
        requestUrl = `${args.options.webUrl}/_api/web/GetFileById(guid'${encodeURIComponent(args.options.id as string)}')`;
      }
      else {
        // concatenate trailing '/' if not provided
        // so if the provided url is for the root site, the substr bellow will get the right value
        let serverRelativeSiteUrl: string = args.options.webUrl;
        if (!serverRelativeSiteUrl.endsWith('/')) {
          serverRelativeSiteUrl = `${serverRelativeSiteUrl}/`;
        }
        serverRelativeSiteUrl = serverRelativeSiteUrl.substr(serverRelativeSiteUrl.indexOf('/', 8));

        let fileUrl: string = args.options.url as string;
        if (!fileUrl.startsWith(serverRelativeSiteUrl)) {
          fileUrl = `${serverRelativeSiteUrl}${fileUrl}`;
        }
        requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`;
      }

      if (args.options.recycle) {
        requestUrl += `/recycle()`;
      }

      const requestOptions: any = {
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
    };

    if (args.options.confirm) {
      await removeFile();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the file ${args.options.id || args.options.url} located in site ${args.options.webUrl}?`
      });

      if (result.continue) {
        await removeFile();
      }
    }
  }
}

module.exports = new SpoFileRemoveCommand();