import { Logger } from '../../cli';
import GlobalOptions from '../../GlobalOptions';
import PeriodBasedReport from './PeriodBasedReport';

interface CommandArgs {
  options: DateAndPeriodBasedOptions;
}

interface DateAndPeriodBasedOptions extends GlobalOptions {
  period?: string;
  date?: string;
}

export default abstract class DateAndPeriodBasedReport extends PeriodBasedReport {
  constructor() {
    super();
    
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        period: args.options.period,
        date: typeof args.options.date !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-d, --date [date]' }
    );

    this.options.forEach(option => {
      option.option = option.option.replace('-p, --period <period>', '-p, --period [period]');
    });
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.period && !args.options.date) {
          return 'Specify period or date, one is required.';
        }
    
        if (args.options.period && args.options.date) {
          return 'Specify period or date but not both.';
        }
    
        if (args.options.date && !((args.options.date as string).match(/^\d{4}-\d{2}-\d{2}$/))) {
          return `${args.options.date} is not a valid date. The supported date format is YYYY-MM-DD`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const periodParameter: string = args.options.period ? `${this.usageEndpoint}(period='${encodeURIComponent(args.options.period)}')` : '';
    const dateParameter: string = args.options.date ? `${this.usageEndpoint}(date=${encodeURIComponent(args.options.date)})` : '';
    const endpoint: string = `${this.resource}/v1.0/reports/${(args.options.period ? periodParameter : dateParameter)}`;
    this.executeReport(endpoint, logger, args.options.output, cb);
  }
}