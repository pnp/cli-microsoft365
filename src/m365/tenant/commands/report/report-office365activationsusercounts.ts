import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class TenantReportOffice365ActivationsUserCountsCommand extends GraphCommand {
  public get name(): string {
    return commands.REPORT_OFFICE365ACTIVATIONSUSERCOUNTS;
  }

  public get description(): string {
    return 'Get the count of enabled users with activated Office subscriptions.';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/reports/getOffice365ActivationsUserCounts`;
    this.loadReport(endpoint, logger, args.options.output, cb);
  }

  private loadReport(endPoint: string, logger: Logger, output: string | undefined, cb: () => void): void {
    const requestOptions: any = {
      url: endPoint,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        let content: string = '';
        const cleanResponse = this.removeEmptyLines(res);

        if (output && output.toLowerCase() === 'json') {
          content = Utils.parseCsvToJson(cleanResponse);
        }
        else {
          content = cleanResponse;
        }

        logger.log(content);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private removeEmptyLines(input: string): string {
    const rows: string[] = input.split('\n');
    const cleanRows = rows.filter(Boolean);
    return cleanRows.join('\n');
  }
}

module.exports = new TenantReportOffice365ActivationsUserCountsCommand();