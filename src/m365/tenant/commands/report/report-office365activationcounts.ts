import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class TenantReportOffice365activationcountsCommand extends GraphCommand {
  public get name(): string {
    return commands.REPORT_OFFICE365ACTIVATIONCOUNTS;
  }

  public get description(): string {
    return 'Get the count of Microsoft 365 activations on desktops and devices.';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/reports/getOffice365ActivationCounts`;
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
          const reportData: any = this.getReport(cleanResponse);
          content = reportData;
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

  private getReport(res: string): any {
    const rows: string[] = res.split('\n');
    const jsonObj: any = [];
    const headers: string[] = rows[0].split(',');

    for (let i = 1; i < rows.length; i++) {
      const data: string[] = rows[i].split(',');
      const obj: any = {};
      for (let j = 0; j < data.length; j++) {
        const value = data[j].trim();
        const numValue = parseInt(value);
        obj[headers[j].trim()] = isNaN(numValue) || numValue.toString() !== value ? value : numValue;
      }
      jsonObj.push(obj);
    }

    return jsonObj;
  }
}

module.exports = new TenantReportOffice365activationcountsCommand();