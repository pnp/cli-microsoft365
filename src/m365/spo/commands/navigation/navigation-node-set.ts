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
  audienceIds?: string;
  id: string;
  isExternal?: boolean;
  title?: string;
  url?: string;
  webUrl: string;
}

class SpoNavigationNodeSetCommand extends SpoCommand {
  public get name(): string {
    return commands.NAVIGATION_NODE_SET;
  }

  public get description(): string {
    return 'Adds a navigation node to the specified site navigation';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        title: typeof args.options.title !== 'undefined',
        url: typeof args.options.url !== 'undefined',
        isExternal: typeof args.options.isExternal !== 'undefined',
        audienceIds: typeof args.options.audienceIds !== 'undefined'
      });
    });
  }

  #initTypes(): void {
    this.types.boolean.push('isExternal');
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--id <id>'
      },
      {
        option: '--title [title]'
      },
      {
        option: '--url [url]'
      },
      {
        option: '--audienceIds [audienceIds]'
      },
      {
        option: '--isExternal [isExternal]'
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

        if (args.options.audienceIds === undefined && !args.options.url && args.options.isExternal === undefined && !args.options.title) {
          return `Please specify atleast one property to update.`;
        }

        if (args.options.audienceIds) {
          const audienceIdsSplitted = args.options.audienceIds.split(',');
          if (audienceIdsSplitted.length > 10) {
            return 'The maximum amount of audienceIds per navigation node exceeded. The maximum amount of audienceIds is 10.';
          }

          if (!validation.isValidGuidArray(audienceIdsSplitted)) {
            return `The option audienceIds contains one or more invalid GUIDs`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Setting navigation node...`);
      }

      let url = args.options.url;
      if (url === '') {
        url = 'http://linkless.header/';
      }

      const requestBody: any = {
        Title: args.options.title,
        IsExternal: args.options.isExternal,
        Url: url
      };

      if (args.options.audienceIds !== undefined) {
        requestBody.AudienceIds = args.options.audienceIds === '' ? [] : args.options.audienceIds.split(',');
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/navigation/GetNodeById(${args.options.id})`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      const response = await request.patch<any>(requestOptions);
      if (response['odata.null'] === true) {
        throw `Navigation node does not exist.`;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

}

module.exports = new SpoNavigationNodeSetCommand();