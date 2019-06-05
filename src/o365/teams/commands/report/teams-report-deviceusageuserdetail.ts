import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from "../../../base/GraphCommand";
import request from '../../../../request';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  period?: string;
  date?: string;
}

class TeamsReportDeviceUsageUserDetailCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_REPORT_DEVICEUSAGEUSERDETAIL}`;
  }

  public get description(): string {
    return 'Gets information about Microsoft Teams device usage by user';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.period = typeof args.options.period !== 'undefined';
    telemetryProps.date = typeof args.options.date !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const periodParameter: string = args.options.period ? `getTeamsDeviceUsageUserDetail(period='${encodeURIComponent(args.options.period)}')` : '';
    const dateParameter: string = args.options.date ? `getTeamsDeviceUsageUserDetail(date=${encodeURIComponent(args.options.date)})` : '';
    const endpoint: string = `${this.resource}/v1.0/reports/${(args.options.period ? periodParameter : dateParameter)}`;

    const requestOptions: any = {
      url: endpoint
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --period [period]',
        description: 'The length of time over which the report is aggregated. Supported values D7|D30|D90|D180',
        autocomplete: ['D7', 'D30', 'D90', 'D180']
      },
      {
        option: '-d, --date [date]',
        description: 'The date for which you would like to view the users who performed any activity. Supported date format is YYYY-MM-DD'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.period && !args.options.date) {
        return 'You can\'t run this command without period or date parameter.';
      }

      if (args.options.period && args.options.date) {
        return 'You can\'t use period and date parameters together.';
      }

      if (args.options.period) {
        if (['D7', 'D30', 'D90', 'D180'].indexOf(args.options.period) < 0) {
          return `${args.options.period} is not a valid period type. The supported values are D7|D30|D90|D180`;
        }
      }

      if (args.options.date && !((args.options.date as string).match(/^\d{4}-\d{2}-\d{2}$/))) {
        return `Provide a valid date in YYYY-MM-DD format`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    As this report is only available for the past 28 days, date parameter value
    should be a date from that range.

  Examples:

    Gets information about Microsoft Teams device usage by user for the last
    week
      ${commands.TEAMS_REPORT_DEVICEUSAGEUSERDETAIL} --period 'D7'

    Gets information about Microsoft Teams device usage by user for
    May 1, 2019
      ${commands.TEAMS_REPORT_DEVICEUSAGEUSERDETAIL} --date 2019-05-01
`);
  }
}

module.exports = new TeamsReportDeviceUsageUserDetailCommand();