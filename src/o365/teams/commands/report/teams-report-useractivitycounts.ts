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
  period: string;
}

class TeamsReportUserActivityCountsCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_REPORT_USERACTIVITYCOUNTS}`;
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams activities by activity type. The activity types are team chat messages, private chat messages, calls, and meetings.';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
      const endpoint: string = `${this.resource}/v1.0/reports/getTeamsUserActivityCounts(period='${encodeURIComponent(args.options.period)}')`;
      
      const requestOptions: any = {
        url: endpoint,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        json: true
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
        option: '-p, --period <period>',
        description: 'The length of time over which the report is aggregated. Supported values D7,D30,D90,D180',
        autocomplete: ['D7', 'D30', 'D90', 'D180']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.period) {
        return 'Required parameter period missing';
      }

      if (['D7', 'D30', 'D90', 'D180'].indexOf(args.options.period) < 0) {
        return `${args.options.period} is not a valid period type. The supported values are D7,D30,D90,D180`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples: 

      Gets the number of Microsoft Teams activities by activity type for last week
      ${commands.TEAMS_REPORT_USERACTIVITYCOUNTS} --period 'D7'
`);
  }
}

module.exports = new TeamsReportUserActivityCountsCommand();