import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: GlobalOptions;
}

class TenantReportOffice365ActivationsUserDetailCommand extends GraphCommand {
  public get name(): string {
    return commands.REPORT_OFFICE365ACTIVATIONSUSERDETAIL;
  }

  public get description(): string {
    return 'Get details about users who have activated Microsoft 365.';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/reports/getOffice365ActivationsUserDetail`;
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
      const cleanResponse = this.removeEmptyLines(res);

      if (output && output.toLowerCase() === 'json') {
        content = formatting.parseCsvToJson(cleanResponse);
      }
      else {
        content = cleanResponse;
      }

      await logger.log(content);

    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private removeEmptyLines(input: string): string {
    const rows: string[] = input.split('\n');
    const cleanRows = rows.filter(Boolean);
    return cleanRows.join('\n');
  }
}

export default new TenantReportOffice365ActivationsUserDetailCommand();