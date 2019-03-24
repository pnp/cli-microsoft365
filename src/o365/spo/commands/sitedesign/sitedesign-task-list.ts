import auth from '../../SpoAuth';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { SiteDesignTask } from './SiteDesignTask';
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): Promise<{ value: SiteDesignTask[] }> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving site designs scheduled for execution on the specified site...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTasks`,
          headers: {
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata',
          },
          json: true
        };

        return request.post(requestOptions);
      })
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
          cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To list information about site designs scheduled for execution on the
    specified site, you have to first log in to a SharePoint site
    using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

  Examples:
  
    List site designs scheduled for execution on the specified site
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/team-a

  More information:

    SharePoint site design and site script overview
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview
`);
  }
}

module.exports = new SpoSiteDesignTaskListCommand();