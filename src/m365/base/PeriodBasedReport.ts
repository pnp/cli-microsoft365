import { Logger } from '../../cli';
import {
  CommandOption
} from '../../Command';
import GlobalOptions from '../../GlobalOptions';
import request from '../../request';
import GraphCommand from "./GraphCommand";

interface CommandArgs {
  options: UsagePeriodOptions;
}

interface UsagePeriodOptions extends GlobalOptions {
  period: string;
}

export default abstract class PeriodBasedReport extends GraphCommand {
  public abstract get usageEndpoint(): string;

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/reports/${this.usageEndpoint}(period='${encodeURIComponent(args.options.period)}')`;
    this.executeReport(endpoint, logger, args.options.output, cb);
  }

  protected executeReport(endPoint: string, logger: Logger, output: string | undefined, cb: () => void): void {
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
        obj[headers[j].trim()] = data[j].trim();
      }
      jsonObj.push(obj);
    }

    return jsonObj;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --period <period>',
        autocomplete: ['D7', 'D30', 'D90', 'D180']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return this.validatePeriod(args.options.period);
  }

  protected validatePeriod(period: string | undefined): boolean | string {
    if (period &&
      ['D7', 'D30', 'D90', 'D180'].indexOf(period) < 0) {
      return `${period} is not a valid period type. The supported values are D7|D30|D90|D180`;
    }

    return true;
  }
}
