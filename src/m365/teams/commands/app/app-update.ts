import request from '../../../../request';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../Utils';
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
          cmd.log(chalk.green('DONE'));
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
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
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
}

module.exports = new TeamsAppUpdateCommand();