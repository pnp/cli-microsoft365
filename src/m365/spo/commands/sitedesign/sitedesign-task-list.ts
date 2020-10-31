import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SiteDesignTask } from './SiteDesignTask';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoSiteDesignTaskListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_TASK_LIST}`;
  }

  public get description(): string {
    return 'Lists site designs scheduled for execution on the specified site';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTasks`,
      headers: {
        accept: 'application/json;odata=nometadata',
      },
      responseType: 'json'
    };

    request.post<{ value: SiteDesignTask[] }>(requestOptions)
      .then((res: { value: SiteDesignTask[] }): void => {
        if (args.options.output === 'json') {
          logger.log(res.value);
        }
        else {
          logger.log(res.value.map(d => {
            return {
              ID: d.ID,
              SiteDesignID: d.SiteDesignID,
              LogonName: d.LogonName
            };
          }));
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
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site for which to list site designs scheduled for execution'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoSiteDesignTaskListCommand();