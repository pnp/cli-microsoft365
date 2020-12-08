import { Logger } from '../../cli';
import {
  CommandOption
} from '../../Command';
import GlobalOptions from '../../GlobalOptions';
import request from '../../request';
import Utils from '../../Utils';
import GraphCommand from "./GraphCommand";

interface CommandArgs {
  options: DateTimeOptions;
}

interface DateTimeOptions extends GlobalOptions {
  fromDateTime: string;
  toDateTime?: string;
}

export default abstract class TeamsCallReport extends GraphCommand {
  public abstract get usageEndpoint(): string;

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/beta/communications/callRecords/${this.usageEndpoint}(fromDateTime=${encodeURIComponent(args.options.fromDateTime)},toDateTime=${encodeURIComponent(args.options.toDateTime!)})`;
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
      .get<{ value: any }>(requestOptions)
      .then((res: { value: any }): void => {
        let content: string = '';

        if (output && output.toLowerCase() === 'json') {
          content = res.value;
        }
        else {
          content = res.value.map((record: { id: string; calleeNumber: string; callerNumber: string; startDateTime: string; }) => ({ id: record.id, calleeNumber: record.calleeNumber, callerNumber: record.callerNumber, startDateTime: record.startDateTime }));
        }

        logger.log(content);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--fromDateTime <fromDateTime>',
        description: 'The start of time range to query. UTC, inclusive'
      },
      {
        option: '--toDateTime [toDateTime]',
        description: 'The end time range to query. UTC, inclusive. Defaults to current DateTime if omitted'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.fromDateTime && !Utils.isValidISODateDashOnly(args.options.fromDateTime)) {
      return 'The fromDateTime is not a valid ISO date string (YYYY-MM-DD)';
    }

    if (args.options.toDateTime && !Utils.isValidISODateDashOnly(args.options.toDateTime)) {
      return 'The toDateTime is not a valid ISO date string (YYYY-MM-DD)';
    }

    if (!args.options.toDateTime) {
      args.options.toDateTime = new Date().toISOString();
    }

    if (Math.ceil((new Date(args.options.toDateTime).getTime() - new Date(args.options.fromDateTime).getTime()) / (1000 * 3600 * 24)) > 90) {
      return 'The maximum number of days between fromDateTime and toDateTime cannot exceed 90'
    }

    return true

  }

}