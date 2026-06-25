import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const behaviorDuringRetentionPeriods = ['doNotRetain', 'retain', 'retainAsRecord', 'retainAsRegulatoryRecord'] as const;
const actionAfterRetentionPeriods = ['none', 'delete', 'startDispositionReview'] as const;
const retentionTriggers = ['dateLabeled', 'dateCreated', 'dateModified', 'dateOfEvent'] as const;
const defaultRecordBehaviors = ['startLocked', 'startUnlocked'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  displayName: z.string().alias('n'),
  behaviorDuringRetentionPeriod: z.enum(behaviorDuringRetentionPeriods),
  actionAfterRetentionPeriod: z.enum(actionAfterRetentionPeriods),
  retentionDuration: z.string().refine(val => !isNaN(Number(val)), {
    message: 'retentionDuration must be a number'
  }),
  retentionTrigger: z.enum(retentionTriggers).optional().alias('t'),
  defaultRecordBehavior: z.enum(defaultRecordBehaviors).optional(),
  descriptionForUsers: z.string().optional(),
  descriptionForAdmins: z.string().optional(),
  labelToBeApplied: z.string().optional(),
  eventTypeId: z.string().optional(),
  eventTypeName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PurviewRetentionLabelAddCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONLABEL_ADD;
  }

  public get description(): string {
    return 'Create a retention label';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => {
        if (opts.retentionTrigger === 'dateOfEvent') {
          return [opts.eventTypeId, opts.eventTypeName].filter(x => x !== undefined).length === 1;
        }
        return true;
      }, {
        message: `Specify either 'eventTypeId' or 'eventTypeName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['eventTypeId', 'eventTypeName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const retentionTrigger: string = args.options.retentionTrigger ? args.options.retentionTrigger : 'dateLabeled';
    const defaultRecordBehavior: string = args.options.defaultRecordBehavior ? args.options.defaultRecordBehavior : 'startLocked';

    const requestBody: any = {
      displayName: args.options.displayName,
      behaviorDuringRetentionPeriod: args.options.behaviorDuringRetentionPeriod,
      actionAfterRetentionPeriod: args.options.actionAfterRetentionPeriod,
      retentionTrigger: retentionTrigger,
      retentionDuration: {
        '@odata.type': '#microsoft.graph.security.retentionDurationInDays',
        days: Number(args.options.retentionDuration)
      },
      defaultRecordBehavior: defaultRecordBehavior
    };

    if (args.options.retentionTrigger === 'dateOfEvent') {
      const eventTypeId = await this.getEventTypeId(args, logger);
      requestBody['retentionEventType@odata.bind'] = `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes/${eventTypeId}`;
    }

    if (args.options.descriptionForAdmins) {
      if (this.verbose) {
        await logger.logToStderr(`Using '${args.options.descriptionForAdmins}' as descriptionForAdmins`);
      }

      requestBody.descriptionForAdmins = args.options.descriptionForAdmins;
    }

    if (args.options.descriptionForUsers) {
      if (this.verbose) {
        await logger.logToStderr(`Using '${args.options.descriptionForUsers}' as descriptionForUsers`);
      }

      requestBody.descriptionForUsers = args.options.descriptionForUsers;
    }

    if (args.options.labelToBeApplied) {
      if (this.verbose) {
        await logger.logToStderr(`Using '${args.options.labelToBeApplied}' as labelToBeApplied...`);
      }

      requestBody.labelToBeApplied = args.options.labelToBeApplied;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/beta/security/labels/retentionLabels`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }

  private async getEventTypeId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.eventTypeId) {
      return args.options.eventTypeId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving the event type id for event type ${args.options.eventTypeName}`);
    }

    const eventTypes = await odata.getAllItems(`${this.resource}/beta/security/triggerTypes/retentionEventTypes`);
    const filteredEventTypes: any = eventTypes.filter((eventType: any) => eventType.displayName === args.options.eventTypeName);

    if (filteredEventTypes.length === 0) {
      throw `The specified retention event type '${args.options.eventTypeName}' does not exist.`;
    }

    return filteredEventTypes[0].id;
  }
}

export default new PurviewRetentionLabelAddCommand();