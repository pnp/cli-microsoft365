import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from "../../../base/GraphCommand";
import request from '../../../../request';
import * as path from 'path';
import * as fs from 'fs';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  period?: string;
  date?: string;
  outputFile?: string;
}

class TeamsReportUserActivityUserDetailCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_REPORT_USERACTIVITYUSERDETAIL}`;
  }

  public get description(): string {
    return 'Get details about Microsoft Teams user activity by user.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.period = typeof args.options.period !== 'undefined';
    telemetryProps.date = typeof args.options.date !== 'undefined';
    telemetryProps.outputFile = typeof args.options.outputFile !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const periodParameter: string = args.options.period ? `getTeamsUserActivityUserDetail(period='${encodeURIComponent(args.options.period)}')` : '';
    const dateParameter: string = args.options.date ? `getTeamsUserActivityUserDetail(date=${encodeURIComponent(args.options.date)})` : '';
    const endpoint: string = `${this.resource}/v1.0/reports/${(args.options.period ? periodParameter : dateParameter)}`;

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

        let content: string = '';

        if(args.options.output && args.options.output.toLowerCase() === 'json')
        {
          let reportdata = this.getJsonReport(res);
          content = JSON.stringify(reportdata);

          if(!args.options.outputFile) {
            cmd.log(reportdata);
          }
        }
        else
        {
          content = res;
          
          if(!args.options.outputFile) {
            cmd.log(content);
          }
        }

        if(args.options.outputFile)
        {
          fs.writeFileSync(args.options.outputFile, content, 'utf8');
          if (this.verbose) {
            cmd.log(`File saved to path '${args.options.outputFile}'`);
          }
        }
        
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getJsonReport(res: string): any {
    let rows = res.split('\n');
    let jsonObj = [];
    let headers = rows[0].split(',');

    for (let i = 1; i < rows.length; i++) {
      let data = rows[i].split(',');
      let obj: any = {};
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
        option: '-p, --period [period]',
        description: 'The length of time over which the report is aggregated. Supported values D7,D30,D90,D180. Specify the period or date, but not both.',
        autocomplete: ['D7', 'D30', 'D90', 'D180']
      },
      {
        option: '-d, --date [date]',
        description: 'The date for which you would like to view the users who performed any activity. Supported date format is YYYY-MM-DD. Specify the date or period, but not both.'
      },
      {
        option: '-f, --outputFile [outputFile]',
        description: 'Path to the file where the upgrade report should be stored in'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.period && !args.options.date) {
        return 'Specify period or date, one is required.';
      }

      if (args.options.period && args.options.date) {
        return 'Specify period or date but not both.';
      }

      if (args.options.period) {
        if (['D7', 'D30', 'D90', 'D180'].indexOf(args.options.period) < 0) {
          return `${args.options.period} is not a valid period type. The supported values are D7,D30,D90,D180`;
        }
      }

      if (args.options.date && !((args.options.date as string).match(/^\d{4}-\d{2}-\d{2}$/))) {
        return `${args.options.date} is not a valid date. The supported date format is YYYY-MM-DD`;
      }

      if (args.options.outputFile && !fs.existsSync(path.dirname(args.options.outputFile))) {
        return 'Specified outputFile path where to save the file does not exist';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

      Gets details about Microsoft Teams user activity by user for 
      the last week
        ${commands.TEAMS_REPORT_USERACTIVITYUSERDETAIL} --period 'D7'

      Gets details about Microsoft Teams user activity by user 
      for July 13, 2019
        ${commands.TEAMS_REPORT_USERACTIVITYUSERDETAIL} --date 2019-07-13

      Gets details about Microsoft Teams user activity by user for the last week
      and exports the report data in the specified path in csv format
        ${commands.TEAMS_REPORT_USERACTIVITYUSERDETAIL} --period D7 --output csv --outputFile 'C:/report.csv'

      Gets details about Microsoft Teams user activity by user for the last week
      and exports the report data in the specified path in json format
        ${commands.TEAMS_REPORT_USERACTIVITYUSERDETAIL} --period D7 --output json --outputFile 'C:/report.json'
`);
  }
}

module.exports = new TeamsReportUserActivityUserDetailCommand();