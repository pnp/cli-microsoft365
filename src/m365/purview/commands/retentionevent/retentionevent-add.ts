import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { odata } from '../../../../utils/odata.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  displayName: z.string().alias('n'),
  eventTypeId: z.string().optional().alias('i'),
  eventTypeName: z.string().optional().alias('e'),
  description: z.string().optional().alias('d'),
  triggerDateTime: z.string().optional()
    .refine(val => val === undefined || validation.isValidISODateTime(val), {
      message: 'The triggerDateTime is not a valid ISO date string'
    }),
  assetIds: z.string().optional().alias('a'),
  keywords: z.string().optional().alias('k')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PurviewRetentionEventAddCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONEVENT_ADD;
  }

  public get description(): string {
    return 'Create a retention event';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.eventTypeId, opts.eventTypeName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'eventTypeId' or 'eventTypeName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['eventTypeId', 'eventTypeName']
        }
      })
      .refine(opts => opts.assetIds !== undefined || opts.keywords !== undefined, {
        error: 'Specify assetIds and/or keywords, but at least one.',
        params: {
          customCode: 'optionSet',
          options: ['assetIds', 'keywords']
        }
      }) as any;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Creating retention event...`);
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
      'retentionEventType@odata.bind': `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes/${eventTypeId}`,
      displayName: args.options.displayName,
      description: args.options.description,
      eventQueries: eventQueries,
      eventTriggerDateTime: args.options.triggerDateTime
    };

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/security/triggers/retentionEvents`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: data
      };

      const res: any = await request.post<any>(requestOptions);
      await logger.log(res);
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
      await logger.logToStderr(`Retrieving the event type id for event type ${args.options.eventTypeName}`);
    }

    const items: any = await odata.getAllItems(`${this.resource}/v1.0/security/triggerTypes/retentionEventTypes`);

    const eventTypes = items.filter((x: any) => x.displayName === args.options.eventTypeName);

    if (eventTypes.length === 0) {
      throw `The specified event type '${args.options.eventTypeName}' does not exist.`;
    }

    return eventTypes[0].id;
  }
}

export default new PurviewRetentionEventAddCommand();