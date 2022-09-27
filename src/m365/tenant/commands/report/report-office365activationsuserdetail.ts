import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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

      logger.log(content);
      
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

module.exports = new TenantReportOffice365ActivationsUserDetailCommand();