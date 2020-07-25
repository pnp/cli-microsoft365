import request from '../../../../request';
import * as fs from 'fs';
import * as path from 'path';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
          cmd.log(chalk.green('DONE'));
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
}

module.exports = new TeamsAppPublishCommand();