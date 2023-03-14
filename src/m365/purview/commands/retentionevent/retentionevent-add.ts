import { Logger } from '../../../../cli/Logger';
import GraphCommand from '../../../base/GraphCommand';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { accessToken } from '../../../../utils/accessToken';
import auth from '../../../../Auth';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import { odata } from '../../../../utils/odata';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  eventTypeId?: string;
  eventTypeName?: string;
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

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --displayName <displayName>'
      },
      {
        option: '-i, --eventTypeId [eventTypeId]'
      },
      {
        option: '-e, --eventTypeName [eventTypeName]'
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

        if (!args.options.assetIds && !args.options.keywords) {
          return 'Specify assetIds and/or keywords, but at least one.';
        }

        return true;
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        description: typeof args.options.description !== 'undefined',
        triggerDateTime: typeof args.options.triggerDateTime !== 'undefined',
        eventTypeId: typeof args.options.eventTypeId !== 'undefined',
        eventTypeName: typeof args.options.eventTypeName !== 'undefined',
        assetIds: typeof args.options.assetIds !== 'undefined',
        keywords: typeof args.options.keywords !== 'undefined'
      });
    });
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['eventTypeId', 'eventTypeName'] }
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

    if (args.options.assetIds) {
      eventQueries.push({ queryType: "files", query: args.options.assetIds });
    }

    if (args.options.keywords) {
      eventQueries.push({ queryType: "messages", query: args.options.keywords });
    }

    const eventTypeId = await this.getEventTypeId(logger, args);

    const data = {
      'retentionEventType@odata.bind': `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes/${eventTypeId}`,
      displayName: args.options.displayName,
      description: args.options.description,
      eventQueries: eventQueries,
      eventTriggerDateTime: args.options.triggerDateTime
    };

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/security/triggers/retentionEvents`,
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

  private async getEventTypeId(logger: Logger, args: CommandArgs): Promise<string> {
    if (args.options.eventTypeId) {
      return args.options.eventTypeId;
    }

    if (this.verbose) {
      logger.logToStderr(`Retrieving the event type id for event type ${args.options.eventTypeName}`);
    }

    const items: any = await odata.getAllItems(`${this.resource}/beta/security/triggers/retentionEvents`);

    const eventTypes = items.filter((x: any) => x.displayName === args.options.eventTypeName);

    if (eventTypes.length === 0) {
      throw `The specified event type '${args.options.eventTypeName}' does not exist.`;
    }

    return eventTypes[0].id;
  }
}

module.exports = new PurviewRetentionEventAddCommand();