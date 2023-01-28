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
        isExternal: args.options.isExternal,
        location: typeof args.options.location !== 'undefined',
        parentNodeId: typeof args.options.parentNodeId !== 'undefined',
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

        if (!args.options.audienceIds && !args.options.url && !args.options.isExternal && !args.options.title) {
          return `Please specify atleast one property to update.`;
        }

        if (args.options.audienceIds) {
          const audienceIdsSplitted = args.options.audienceIds.split(',');
          if (audienceIdsSplitted.length > 10) {
            return 'The maximum amount of audienceIds per navigation node exceeded. The maximum amount of auciendeIds to be set is 10.';
          }
          const invalidAudienceIds: string[] = [];
          audienceIdsSplitted.map(audienceId => {
            if (!validation.isValidGuid(audienceId)) {
              invalidAudienceIds.push(audienceId);
            }
          });
          if (invalidAudienceIds.length > 0) {
            return `Invalid audienceIds have been entered. Invalid ids are: ${invalidAudienceIds.join(',')}.`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Setting navigation node...`);
    }
    logger.log(args.options);
    const requestBody = this.mapRequestBody(args.options);

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/navigation/GetNodeById(${args.options.id})`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};
    if (options.title) {
      requestBody.Title = options.title;
    }
    if (options.isExternal !== undefined) {
      requestBody.IsExternal = options.isExternal;
    }
    if (options.url) {
      requestBody.Url = options.url;
    }
    if (options.audienceIds !== undefined) {
      if (options.audienceIds === '') {
        requestBody.AudienceIds = [];
      }
      else {
        requestBody.AudienceIds = options.audienceIds.split(',');
      }

    }
    return requestBody;
  }
}

module.exports = new SpoNavigationNodeSetCommand();