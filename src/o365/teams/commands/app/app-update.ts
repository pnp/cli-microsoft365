import request from '../../../../request';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../Utils';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  filePath: string;
}

class TeamsAppUpdateCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_APP_UPDATE}`;
  }

  public get description(): string {
    return 'Updates Teams app in the organization\'s app catalog';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const { id: appId, filePath } = args.options;

    const fullPath: string = path.resolve(filePath);
    if (this.verbose) {
      cmd.log(`Updating app with id '${appId}' and file '${fullPath}' in the app catalog...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/appCatalogs/teamsApps/${appId}`,
      headers: {
        "content-type": "application/zip"
      },
      body: fs.readFileSync(fullPath)
    };

    request
      .put(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (res: any): void => this.handleRejectedODataJsonPromise(res, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the app to update'
      },
      {
        option: '-p, --filePath <filePath>',
        description: 'Absolute or relative path to the Teams manifest zip file to update in the app catalog'
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

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (!args.options.filePath) {
        return 'Missing required option filePath';
      }

      const fullPath: string = path.resolve(args.options.filePath);

      if (!fs.existsSync(fullPath)) {
        return `File '${fullPath}' not found`;
      }

      if (fs.lstatSync(fullPath).isDirectory()) {
        return `Path '${fullPath}' points to a directory`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    You can only update a Teams app as a global administrator.

  Examples:

    Update the Teams app with ID ${chalk.grey('83cece1e-938d-44a1-8b86-918cf6151957')}
    from file ${chalk.grey('teams-manifest.zip')}
      ${this.name} --id 83cece1e-938d-44a1-8b86-918cf6151957 --filePath ./teams-manifest.zip
`);
  }
}

module.exports = new TeamsAppUpdateCommand();