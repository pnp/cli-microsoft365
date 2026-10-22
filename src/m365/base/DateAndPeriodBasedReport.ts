import { z } from 'zod';
import { Logger } from '../../cli/Logger.js';
import { globalOptionsZod } from '../../Command.js';
import { formatting } from '../../utils/formatting.js';
import PeriodBasedReport from './PeriodBasedReport.js';

export const dateAndPeriodBasedReportOptions = z.strictObject({
  ...globalOptionsZod.shape,
  output: z.enum(['json', 'csv']).optional().alias('o'),
  period: z.enum(['D7', 'D30', 'D90', 'D180']).optional().alias('p'),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/, 'The supported date format is YYYY-MM-DD').optional().alias('d')
}).refine(opts => opts.period || opts.date, {
  message: `Specify either 'period' or 'date'.`,
  params: {
    customCode: 'optionSet',
    options: ['period', 'date']
  }
}).refine(opts => !(opts.period && opts.date), {
  message: `Specify either 'period' or 'date', but not both.`,
  params: {
    customCode: 'optionSet',
    options: ['period', 'date']
  }
});

declare type Options = z.infer<typeof dateAndPeriodBasedReportOptions>;

export interface CommandArgs {
  options: Options;
}

export default abstract class DateAndPeriodBasedReport extends PeriodBasedReport {
  public get schema(): z.ZodType | undefined {
    return dateAndPeriodBasedReportOptions;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const periodParameter: string = args.options.period ? `${this.usageEndpoint}(period='${formatting.encodeQueryParameter(args.options.period)}')` : '';
    const dateParameter: string = args.options.date ? `${this.usageEndpoint}(date=${formatting.encodeQueryParameter(args.options.date)})` : '';
    const endpoint: string = `${this.resource}/v1.0/reports/${(args.options.period ? periodParameter : dateParameter)}`;
    await this.executeReport(endpoint, logger, args.options.output);
  }
}