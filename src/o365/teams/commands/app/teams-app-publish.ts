import request from '../../../../request';
import * as fs from 'fs';
import * as path from 'path';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filePath: string;
}

class TeamsAppPublishCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_APP_PUBLISH}`;
  }

  public get description(): string {
    return 'Publishes Teams app to the organization\'s app catalog';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const fullPath: string = path.resolve(args.options.filePath);
    if (this.verbose) {
      cmd.log(`Adding app '${fullPath}' to app catalog...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/appCatalogs/teamsApps`,
      headers: {
        "content-type": "application/zip",
        accept: 'application/json;odata.metadata=none'
      },
      body: fs.readFileSync(fullPath)
    };

    request
      .post<{ id: string; }>(requestOptions)
      .then((res: { id: string; }): void => {
        if (res && res.id) {
          cmd.log(res.id);
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (res: any): void => this.handleRejectedODataJsonPromise(res, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --filePath <filePath>',
        description: 'Absolute or relative path to the Teams manifest zip file to add to the app catalog'
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

    You can only publish a Teams app as a global administrator.

  Examples:

    Add the ${chalk.grey('teams-manifest.zip')} file to the organization's app catalog
      ${this.name} --filePath ./teams-manifest.zip
`);
  }
}

module.exports = new TeamsAppPublishCommand();