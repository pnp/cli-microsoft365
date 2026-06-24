import { z } from 'zod';
import { Logger } from '../../cli/Logger.js';
import { globalOptionsZod } from '../../Command.js';
import request, { CliRequestOptions } from '../../request.js';
import { formatting } from '../../utils/formatting.js';
import GraphCommand from "./GraphCommand.js";

export const periodBasedReportOptions = z.strictObject({
  ...globalOptionsZod.shape,
  output: z.enum(['json', 'csv']).optional().alias('o'),
  period: z.enum(['D7', 'D30', 'D90', 'D180']).alias('p')
});

declare type Options = z.infer<typeof periodBasedReportOptions>;

export interface CommandArgs {
  options: Options;
}

export default abstract class PeriodBasedReport extends GraphCommand {
  public abstract get usageEndpoint(): string;

  protected get allowedOutputs(): string[] {
    return ['json', 'csv'];
  }

  public get schema(): z.ZodType | undefined {
    return periodBasedReportOptions;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/reports/${this.usageEndpoint}(period='${formatting.encodeQueryParameter(args.options.period)}')`;
    await this.executeReport(endpoint, logger, args.options.output);
  }

  protected async executeReport(endPoint: string, logger: Logger, output: string | undefined): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: endPoint,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    let res: any;
    try {
      res = await request.get(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
      return;
    }

    let content: string;
    const cleanResponse = this.removeEmptyLines(res);

    if (output && output.toLowerCase() === 'json') {
      const reportData: any = this.getReport(cleanResponse);
      content = reportData;
    }
    else {
      content = cleanResponse;
    }

    await logger.log(content);
  }

  private removeEmptyLines(input: string): string {
    const rows: string[] = input.split('\n');
    const cleanRows = rows.filter(Boolean);
    return cleanRows.join('\n');
  }

  private getReport(res: string): any {
    const rows: string[] = res.split('\n');
    const jsonObj: any = [];
    const headers: string[] = rows[0].split(',');

    for (let i = 1; i < rows.length; i++) {
      const data: string[] = rows[i].split(',');
      const obj: any = {};
      for (let j = 0; j < data.length; j++) {
        obj[headers[j].trim()] = data[j].trim();
      }
      jsonObj.push(obj);
    }

    return jsonObj;
  }
}
