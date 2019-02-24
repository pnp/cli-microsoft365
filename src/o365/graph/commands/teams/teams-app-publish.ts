import * as request from 'request-promise-native';
import * as fs from 'fs';
import * as path from 'path';
import auth from '../../GraphAuth';
import Utils from '../../../../Utils';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import { Team } from './Team';
import { GraphItemsListCommand } from '../GraphItemsListCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filePath: string;
}

class TeamsAppPublishCommand extends GraphItemsListCommand<Team> {
  public get name(): string {
    return `${commands.TEAMS_APP_PUBLISH}`;
  }

  public get description(): string {
    return 'Publish Teams app to your organization\'s app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let endpoint: string = `${auth.service.resource}/v1.0/appCatalogs/teamsApps`;
    
    auth.ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): request.RequestPromise => {
          const fullPath: string = path.resolve(args.options.filePath);
          if (this.verbose) {
            cmd.log(`Adding app '${fullPath}' to app catalog...`);
          }

          const requestOptions: any = {
            url: endpoint,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              "content-type": "application/zip",
              accept: 'application/json;odata.metadata=none'
            }),
            body: fs.readFileSync(fullPath)
          };
  
          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }
  
          return request.post(requestOptions);
        })
        .then((res: { id: string; }): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }
  
          if (res && res.id) {
            cmd.log(res.id);
          }
  
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --filePath <filePath>',
        description: 'Absolute or relative path to the Teams Manifest zip file to add to the app catalog'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.filePath) {
        return 'Missing required option filePath';
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

    To publish Microsoft Teams apps, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    Add the ${chalk.grey('teams-manifest.zip')} file to the organization's app catalog
      ${chalk.grey(config.delimiter)} ${this.name} --filePath ./teams-manifest.zip
`);
  }
}

module.exports = new TeamsAppPublishCommand();