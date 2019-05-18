import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../GraphItemsListCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

const maxDays = 30;
const maxDate: Date = new Date();
maxDate.setDate(maxDate.getDate() - maxDays);

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  period?: string;
  date?: string;
}

class GraphReportsTeamsDeviceUsageUserDetailCommand extends GraphItemsListCommand<any> {
  public get name(): string {
    return `${commands.REPORTS_TEAMS_DEVICE_USAGE_USER_DETAIL}`;
  }

  public get description(): string {
    return 'Gets detail about Microsoft Teams device usage by user';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.period = typeof args.options.period !== 'undefined';
    telemetryProps.date = typeof args.options.date !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const periodParameter: string = args.options.period ? `getTeamsDeviceUsageUserDetail(period='${encodeURIComponent(args.options.period)}')` : '';
    const dateParameter: string = args.options.date ? `getTeamsDeviceUsageUserDetail(date='${encodeURIComponent(args.options.date)}')` : '';

    let endpoint: string = '';

    if (args.options.period) {
      endpoint = `${auth.service.resource}/v1.0/reports/${periodParameter}`;
    }

    if (args.options.date) {
      endpoint = `${auth.service.resource}/v1.0/reports/${dateParameter}`;
    }

    console.log(endpoint);

  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --period [period]',
        description: 'Specify the length of time over which the report is aggregated. The supported values are D7|D30|D90|D180',
        autocomplete: ['D7', 'D30', 'D90', 'D180']
      },
      {
        option: '-d, --date [date]',
        description: 'Specify the date for which you would like to view the users who performed any activity. The supported date format is YYYY-MM-DD'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {

    return (args: CommandArgs): boolean | string => {

      if (args.options.period && args.options.date) {
        return 'You can\'t use period and date parameters together.';
      }

      if (args.options.period) {
        if (args.options.period !== 'D7' && args.options.period !== 'D30' && args.options.period !== 'D90' && args.options.period !== 'D180') {
          return `${args.options.period} is not a valid period type. The supported values are D7|D30|D90|D180`;
        }
      }

      const parsedDate = Date.parse(args.options.date as string);

      if (args.options.date && !(!parsedDate) !== true) {
        return `Provide the date in YYYY-MM-DD format`;
      }

      if (new Date(parsedDate) <= maxDate) {
        return `This report is only available for the past 30 days, date value should be a date from that range.`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To get details about Microsoft Teams device usage by user, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.
    
    Reports.Read.All permission is required to call this API.

  Examples:

  Gets detail about Microsoft Teams device usage by user for the length of time over which 
  the report is aggregated
    ${chalk.grey(config.delimiter)} ${commands.REPORTS_TEAMS_DEVICE_USAGE_USER_DETAIL} --period D7

  Gets detail about Microsoft Teams device usage by user for date for which you would like to 
  view the users who performed any activity
    ${chalk.grey(config.delimiter)} ${commands.REPORTS_TEAMS_DEVICE_USAGE_USER_DETAIL} --date 2019-05-01
`);
  }
}

module.exports = new GraphReportsTeamsDeviceUsageUserDetailCommand();