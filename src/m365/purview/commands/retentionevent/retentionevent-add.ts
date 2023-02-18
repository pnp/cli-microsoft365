import { Logger } from '../../../../cli/Logger';
import GraphCommand from '../../../base/GraphCommand';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { accessToken } from '../../../../utils/accessToken';
import auth from '../../../../Auth';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  eventType: string;
  description?: string;
  triggerDateTime?: string;
  assetIds?: string;
  keywords?: string;
}

class PurviewRetentionEventAddCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONEVENT_ADD;
  }

  public get description(): string {
    return 'Create a retention event';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --displayName <displayName>'
      },
      {
        option: '-t, --eventType <eventType>'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--triggerDateTime [triggerDateTime]'
      },
      {
        option: '-a, --assetIds [assetIds]'
      },
      {
        option: '-k, --keywords [keywords]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.triggerDateTime && !validation.isValidISODateTime(args.options.triggerDateTime)) {
          return 'The triggerDateTime is not a valid ISO date string';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.');
    }

    if (this.verbose) {
      logger.logToStderr(`Creating retention event...`);
    }

    const eventQueries: any[] = [];

    args.options.assetIds?.split(',').forEach(x => { eventQueries.push({ queryType: "files", query: x }); });
    args.options.keywords?.split(',').forEach(x => { eventQueries.push({ queryType: "messages", query: x }); });

    const data = {
      retentionEventType: args.options.eventType,
      displayName: args.options.displayName,
      description: args.options.description,
      eventQueries: eventQueries,
      eventTriggerDateTime: args.options.triggerDateTime
    };

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/security/triggers/retentionEvents/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: data
      };

      const res: any = await request.post<any>(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PurviewRetentionEventAddCommand();