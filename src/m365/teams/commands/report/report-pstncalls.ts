import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { TeamsPstnCall } from './TeamsPstnCall';

interface CommandArgs {
  options: DateTimeOptions;
}

interface DateTimeOptions extends GlobalOptions {
  fromDateTime: string;
  toDateTime?: string;
}
class TeamsReportPstnCallsCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_REPORT_PSTNCALLS}`;
  }

  public get description(): string {
    return 'Get details about PSTN calls made within a given time period';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'calleeNumber', 'callerNumber', 'startDateTime'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const toDateTimeParameter: string = encodeURIComponent(args.options.toDateTime ? args.options.toDateTime : new Date().toISOString());
    const endpoint: string = `${this.resource}/beta/communications/callRecords/getPstnCalls(fromDateTime=${encodeURIComponent(args.options.fromDateTime)},toDateTime=${toDateTimeParameter})`;
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
      .get<{ value: TeamsPstnCall[] }>(requestOptions)
      .then((res: { value: TeamsPstnCall[] }): void => {

        logger.log(res);

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
        description: 'The end time range to query. UTC, inclusive. Defaults to today if omitted'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidISODateDashOnly(args.options.fromDateTime)) {
      return 'The fromDateTime is not a valid ISO date string (YYYY-MM-DD)';
    }

    if (args.options.toDateTime && !Utils.isValidISODateDashOnly(args.options.toDateTime)) {
      return 'The toDateTime is not a valid ISO date string (YYYY-MM-DD)';
    }

    if (Math.ceil((new Date(args.options.toDateTime || new Date().toISOString()).getTime() - new Date(args.options.fromDateTime).getTime()) / (1000 * 3600 * 24)) > 90) {
      return 'The maximum number of days between fromDateTime and toDateTime cannot exceed 90'
    }

    return true

  }

}

module.exports = new TeamsReportPstnCallsCommand();