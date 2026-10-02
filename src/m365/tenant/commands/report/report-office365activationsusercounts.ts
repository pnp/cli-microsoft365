import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = globalOptionsZod.strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TenantReportOffice365ActivationsUserCountsCommand extends GraphCommand {
  protected get allowedOutputs(): string[] {
    return ['json', 'csv'];
  }

  public get name(): string {
    return commands.REPORT_OFFICE365ACTIVATIONSUSERCOUNTS;
  }

  public get description(): string {
    return 'Gets the count of users that are enabled and those that have activated the Office subscription on desktop or devices or shared computers';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/reports/getOffice365ActivationsUserCounts`;
    await this.loadReport(endpoint, logger, args.options.output);
  }

  private async loadReport(endPoint: string, logger: Logger, output: string | undefined): Promise<void> {
    const requestOptions: any = {
      url: endPoint,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res: any = await request.get(requestOptions);
      let content: string = '';

      if (output && output.toLowerCase() === 'json') {
        content = formatting.parseCsvToJson(res);
      }
      else {
        content = res;
      }

      await logger.log(content);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

}

export default new TenantReportOffice365ActivationsUserCountsCommand();