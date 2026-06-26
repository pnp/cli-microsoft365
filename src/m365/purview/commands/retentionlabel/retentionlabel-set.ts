import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const behaviorDuringRetentionPeriodValues = ['doNotRetain', 'retain', 'retainAsRecord', 'retainAsRegulatoryRecord'] as const;
const actionAfterRetentionPeriodValues = ['none', 'delete', 'startDispositionReview'] as const;
const retentionTriggerValues = ['dateLabeled', 'dateCreated', 'dateModified', 'dateOfEvent'] as const;
const defaultRecordBehaviorValues = ['startLocked', 'startUnlocked'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), {
    error: 'The value must be a valid GUID.'
  }).alias('i'),
  behaviorDuringRetentionPeriod: z.enum(behaviorDuringRetentionPeriodValues).optional(),
  actionAfterRetentionPeriod: z.enum(actionAfterRetentionPeriodValues).optional(),
  retentionDuration: z.string().refine(val => !isNaN(Number(val)), {
    error: 'retentionDuration must be a number'
  }).optional(),
  retentionTrigger: z.enum(retentionTriggerValues).optional().alias('t'),
  defaultRecordBehavior: z.enum(defaultRecordBehaviorValues).optional(),
  descriptionForUsers: z.string().optional(),
  descriptionForAdmins: z.string().optional(),
  labelToBeApplied: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PurviewRetentionLabelSetCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONLABEL_SET;
  }

  public get description(): string {
    return 'Update a retention label';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.behaviorDuringRetentionPeriod || opts.actionAfterRetentionPeriod || opts.retentionDuration || opts.retentionTrigger || opts.defaultRecordBehavior || opts.descriptionForUsers || opts.descriptionForAdmins || opts.labelToBeApplied, {
        error: 'Specify at least one property to update.',
        path: ['behaviorDuringRetentionPeriod'],
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.log(`Starting to update retention label with id ${args.options.id}`);
    }

    const requestBody = this.mapRequestBody(args.options);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/beta/security/labels/retentionLabels/${args.options.id}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json',
      data: requestBody
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
    const excludeOptions: string[] = [
      'debug',
      'verbose',
      'output',
      'id',
      'retentionDuration'
    ];
    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        requestBody[key] = `${(<any>options)[key]}`;
      }
    });

    if (options.retentionDuration) {
      requestBody['retentionDuration'] = {
        '@odata.type': 'microsoft.graph.security.retentionDurationInDays',
        'days': Number(options.retentionDuration)
      };
    }
    return requestBody;
  }
}

export default new PurviewRetentionLabelSetCommand();