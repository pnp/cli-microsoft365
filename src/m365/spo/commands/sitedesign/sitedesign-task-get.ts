import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SiteDesignTask } from './SiteDesignTask';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getSpoUrl(logger, this.debug)
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
          logger.log(res);
        }

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.taskId)) {
      return `${args.options.taskId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoSiteDesignTaskGetCommand();