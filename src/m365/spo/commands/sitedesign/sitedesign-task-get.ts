import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { SiteDesignTask } from './SiteDesignTask';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  taskId: string;
}

class SpoSiteDesignTaskGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_TASK_GET}`;
  }

  public get description(): string {
    return 'Gets information about the specified site design scheduled for execution';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getSpoUrl(cmd, this.debug)
      .then((spoUrl: string): Promise<SiteDesignTask> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTask`,
          headers: {
            accept: 'application/json;odata=nometadata',
          },
          body: {
            taskId: args.options.taskId
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((res: SiteDesignTask): void => {
        if (!res["odata.null"]) {
          cmd.log(res);
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --taskId <taskId>',
        description: 'The ID of the site design task to get information for'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.taskId)) {
        return `${args.options.taskId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new SpoSiteDesignTaskGetCommand();