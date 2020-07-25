import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { SiteDesignTask } from './SiteDesignTask';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTasks`,
      headers: {
        accept: 'application/json;odata=nometadata',
      },
      json: true
    };

    request.post<{ value: SiteDesignTask[] }>(requestOptions)
      .then((res: { value: SiteDesignTask[] }): void => {
        if (args.options.output === 'json') {
          cmd.log(res.value);
        }
        else {
          cmd.log(res.value.map(d => {
            return {
              ID: d.ID,
              SiteDesignID: d.SiteDesignID,
              LogonName: d.LogonName
            };
          }));
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
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site for which to list site designs scheduled for execution'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoSiteDesignTaskListCommand();