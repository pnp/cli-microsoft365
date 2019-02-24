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
  id: string;
  filePath: string;
}

class TeamsAppUpdateCommand extends GraphItemsListCommand<Team> {
  public get name(): string {
    return `${commands.TEAMS_APP_UPDATE}`;
  }

  public get description(): string {
    return 'Update a Teams app in your organization\'s app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const { id: appId, filePath } = args.options;
    let endpoint: string = `${auth.service.resource}/v1.0/appCatalogs/teamsApps/${appId}`;
    
    auth.ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): request.RequestPromise => {
          const fullPath: string = path.resolve(filePath);
          if (this.verbose) {
            cmd.log(`Updating app with id '${appId}' and file '${fullPath}' in your app catalog...`);
          }

          const requestOptions: any = {
            url: endpoint,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              "content-type": "application/zip"
            }),
            body: fs.readFileSync(fullPath)
          };
  
          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }
  
          return request.put(requestOptions);
        })
        .then((res: string): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
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
        option: '-i, --id <id>',
        description: 'ID of the app to upgrade'
      },
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
      if (!args.options.id) {
        return 'Missing required option id';
      }

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

    To update Microsoft Teams apps, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    Update the Teams app with ID ${chalk.grey('83cece1e-938d-44a1-8b86-918cf6151957')}
      ${chalk.grey(config.delimiter)} ${this.name} --filePath ./teams-manifest.zip
`);
  }
}

module.exports = new TeamsAppUpdateCommand();