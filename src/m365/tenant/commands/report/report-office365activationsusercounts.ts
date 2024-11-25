import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: GlobalOptions;
}

class TenantReportOffice365ActivationsUserCountsCommand extends GraphCommand {
  protected get allowedOutputs(): string[] {
    return ['json', 'csv'];
  }

  public get name(): string {
    return commands.REPORT_OFFICE365ACTIVATIONSUSERCOUNTS;
  }

  public get description(): string {
    return 'Get the count of enabled users with activated Office subscriptions.';
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