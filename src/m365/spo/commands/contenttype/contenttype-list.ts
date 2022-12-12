import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  category?: string;
}

class SpoContentTypeListCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_LIST;
  }

  public get description(): string {
    return 'Lists all available content types in the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['StringId', 'Name', 'Hidden', 'ReadOnly', 'Sealed'];
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
        category: typeof args.options.category !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-c, --category [category]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let requestUrl: string = `${args.options.webUrl}/_api/web/ContentTypes`;

      if (args.options.category) {
        requestUrl += `?$filter=Group eq '${formatting.encodeQueryParameter(args.options.category as string)}'`;
      }

      const res = await odata.getAllItems<any>(requestUrl);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoContentTypeListCommand();