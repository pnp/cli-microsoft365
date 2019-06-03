import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from "../../GraphCommand";
import request from '../../../../request';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  period?: string;
  date?: string;
}

class GraphReportTeamsDeviceUsageUserDetailCommand extends GraphCommand {

  public get name(): string {
    return `${commands.REPORT_TEAMSDEVICEUSAGEUSERDETAIL}`;
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
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): Promise<{}> => {

        const periodParameter: string = args.options.period ? `getTeamsDeviceUsageUserDetail(period='${encodeURIComponent(args.options.period)}')` : '';
        const dateParameter: string = args.options.date ? `getTeamsDeviceUsageUserDetail(date=${encodeURIComponent(args.options.date)})` : '';
        const endpoint: string = `${auth.service.resource}/v1.0/reports/${(args.options.period ? periodParameter : dateParameter)}`;

        const requestOptions: any = {
          url: endpoint,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`
          }
        };

        return request.get(requestOptions);
      })
      .then((res: any): void => {
        cmd.log(res);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
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

      if (!args.options.period && !args.options.date) {
        return 'You can\'t run this command without period or date parameter.';
      }

      if (args.options.period && args.options.date) {
        return 'You can\'t use period and date parameters together.';
      }

      if (args.options.period) {
        const period: string = args.options.period;
        if (period !== 'D7' && period !== 'D30' && period !== 'D90' && period !== 'D180' && period !== 'd7' && period !== 'd30' && period !== 'd90' && period !== 'd180') {
          return `${period} is not a valid period type. The supported values are D7|D30|D90|D180`;
        }
      }

      const maxDays: number = 30;
      const maxDate: Date = new Date();
      maxDate.setDate(maxDate.getDate() - maxDays);
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
    
    As this report is only available for the past 30 days, date parameter value should be a date from that range.

  Examples:

  Gets detail about Microsoft Teams device usage by user for the length of time over which 
  the report is aggregated
    ${chalk.grey(config.delimiter)} ${commands.REPORT_TEAMSDEVICEUSAGEUSERDETAIL} --period 'D7'

  Gets detail about Microsoft Teams device usage by user for date for which you would like to 
  view the users who performed any activity
    ${chalk.grey(config.delimiter)} ${commands.REPORT_TEAMSDEVICEUSAGEUSERDETAIL} --date 2019-05-01
`);
  }
}

module.exports = new GraphReportTeamsDeviceUsageUserDetailCommand();