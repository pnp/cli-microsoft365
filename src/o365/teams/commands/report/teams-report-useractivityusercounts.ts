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
  period: string;
  outputFile?: string;
}

class TeamsReportUserActivityUserCountsCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_REPORT_USERACTIVITYUSERCOUNTS}`;
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams users by activity type. The activity types are number of teams chat messages, private chat messages, calls, or meetings.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.period = args.options.period;
    telemetryProps.outputFile = typeof args.options.outputFile !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/reports/getTeamsUserActivityUserCounts(period='${encodeURIComponent(args.options.period)}')`;

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
        let cleanResponse = this.removeEmptyLines(res);

        if (args.options.output && args.options.output.toLowerCase() === 'json') {
          const reportdata: any = this.getReport(cleanResponse);
          content = JSON.stringify(reportdata);
        }
        else {
          content = cleanResponse;
        }

        if (!args.options.outputFile) {
          cmd.log(content);
        }
        else {
          fs.writeFileSync(args.options.outputFile, content, 'utf8');
          if (this.verbose) {
            cmd.log(`File saved to path '${args.options.outputFile}'`);
          }
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
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
        option: '-p, --period <period>',
        description: 'The length of time over which the report is aggregated. Supported values D7|D30|D90|D180',
        autocomplete: ['D7', 'D30', 'D90', 'D180']
      },
      {
        option: '-f, --outputFile [outputFile]',
        description: 'Path to the file where the Microsoft Teams users by activity type report should be stored in'
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
        return `${args.options.period} is not a valid period type. The supported values are D7|D30|D90|D180`;
      }

      if (args.options.outputFile && !fs.existsSync(path.dirname(args.options.outputFile))) {
        return `The specified path ${path.dirname(args.options.outputFile)} doesn't exist`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples: 

    Gets the number of Microsoft Teams users by activity type for last week
      ${commands.TEAMS_REPORT_USERACTIVITYUSERCOUNTS} --period D7

    Gets the number of Microsoft Teams users by activity type for last week
    and exports the report data in the specified path in text format
      ${commands.TEAMS_REPORT_USERACTIVITYUSERCOUNTS} --period D7 --output text --outputFile 'C:/report.txt'

    Gets the number of Microsoft Teams users by activity type for last week
    and exports the report data in the specified path in json format
      ${commands.TEAMS_REPORT_USERACTIVITYUSERCOUNTS} --period D7 --output json --outputFile 'C:/report.json'
`);
  }
}

module.exports = new TeamsReportUserActivityUserCountsCommand();