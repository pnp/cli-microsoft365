import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphApplicationCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  output: z.enum(['csv', 'json']).optional().alias('o'),
  fromDateTime: z.string().refine(val => validation.isValidISODateDashOnly(val), {
    message: 'The fromDateTime is not a valid ISO date string (YYYY-MM-DD)'
  }),
  toDateTime: z.string().refine(val => validation.isValidISODateDashOnly(val), {
    message: 'The toDateTime is not a valid ISO date string (YYYY-MM-DD)'
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsReportDirectroutingcallsCommand extends GraphApplicationCommand {
  protected get allowedOutputs(): string[] {
    return ['json', 'csv'];
  }

  public get name(): string {
    return commands.REPORT_DIRECTROUTINGCALLS;
  }

  public get description(): string {
    return 'Gets details about direct routing calls made within a given time period';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'calleeNumber', 'callerNumber', 'startDateTime'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => Math.ceil((new Date(opts.toDateTime || new Date().toISOString()).getTime() - new Date(opts.fromDateTime).getTime()) / (1000 * 3600 * 24)) <= 90, {
      message: 'The maximum number of days between fromDateTime and toDateTime cannot exceed 90'
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const toDateTimeParameter: string = formatting.encodeQueryParameter(args.options.toDateTime ? args.options.toDateTime : new Date().toISOString());

    try {
      const response = await odata.getAllItems<any>(`${this.resource}/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=${formatting.encodeQueryParameter(args.options.fromDateTime)},toDateTime=${toDateTimeParameter})`);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsReportDirectroutingcallsCommand();