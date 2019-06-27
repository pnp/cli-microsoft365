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
  period: string;
}

class GraphReportTeamsDeviceUsageUserCountsCommand extends GraphCommand {
  public get name(): string {
    return `${commands.REPORT_TEAMSDEVICEUSAGEUSERCOUNTS}`;
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams daily unique users by device type';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): Promise<{}> => {
        const endpoint: string = `${auth.service.resource}/v1.0/reports/getTeamsDeviceUsageUserCounts(period='${encodeURIComponent(args.options.period)}')`;

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
        option: '-p, --period <period>',
        description: 'The length of time over which the report is aggregated. Supported values D7|D30|D90|D180',
        autocomplete: ['D7', 'D30', 'D90', 'D180']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.period) {
        return 'You can\'t run this command without period parameter.';
      }

      if (['D7', 'D30', 'D90', 'D180'].indexOf(args.options.period) < 0) {
        return `${args.options.period} is not a valid period type. The supported values are D7|D30|D90|D180`;
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

    To get the number of Microsoft Teams daily unique users by device type, you have to first
    log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    Gets the number of Microsoft Teams daily unique users by device type for the last week
      ${chalk.grey(config.delimiter)} ${commands.REPORT_TEAMSDEVICEUSAGEUSERCOUNTS} --period 'D7'
`);
  }
}

module.exports = new GraphReportTeamsDeviceUsageUserCountsCommand();