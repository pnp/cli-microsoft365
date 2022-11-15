import { Logger } from '../../cli/Logger';
import GlobalOptions from '../../GlobalOptions';
import request from '../../request';
import { formatting } from '../../utils/formatting';
import GraphCommand from "./GraphCommand";

interface CommandArgs {
  options: UsagePeriodOptions;
}

interface UsagePeriodOptions extends GlobalOptions {
  period: string;
}

export default abstract class PeriodBasedReport extends GraphCommand {
  public abstract get usageEndpoint(): string;

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.push(
      {
        option: '-p, --period <period>',
        autocomplete: ['D7', 'D30', 'D90', 'D180']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      (args) => this.validatePeriod(args),
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/reports/${this.usageEndpoint}(period='${formatting.encodeQueryParameter(args.options.period)}')`;
    await this.executeReport(endpoint, logger, args.options.output);
  }

  protected async executeReport(endPoint: string, logger: Logger, output: string | undefined): Promise<void> {
    const requestOptions: any = {
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

  protected async validatePeriod(args: CommandArgs): Promise<boolean | string> {
    const period = args.options.period;
    if (period &&
      ['D7', 'D30', 'D90', 'D180'].indexOf(period) < 0) {
      return `${period} is not a valid period type. The supported values are D7|D30|D90|D180`;
    }

    return true;
  }
}
