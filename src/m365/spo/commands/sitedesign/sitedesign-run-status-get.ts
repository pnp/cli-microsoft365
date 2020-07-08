import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { SiteScriptActionStatus } from './SiteScriptActionStatus';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  runId: string;
  webUrl: string;
}

class SpoSiteDesignRunStatusGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_RUN_STATUS_GET}`;
  }

  public get description(): string {
    return 'Gets information about the site scripts executed for the specified site design';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const body: any = {
      runId: args.options.runId
    };

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata'
      },
      body: body,
      json: true
    };

    request.post<{ value: SiteScriptActionStatus[] }>(requestOptions)
      .then((res: { value: SiteScriptActionStatus[] }): void => {
        if (args.options.output === 'json') {
          cmd.log(res.value);
        }
        else {
          cmd.log(res.value.map(s => {
            return {
              ActionTitle: s.ActionTitle,
              SiteScriptTitle: s.SiteScriptTitle,
              OutcomeText: s.OutcomeText
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
        description: 'The URL of the site for which to get the information'
      },
      {
        option: '-i, --runId <runId>',
        description: 'ID of the site design applied to the site as retrieved using \'spo sitedesign run list\''
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

      if (!args.options.runId) {
        return 'Required parameter runId missing';
      }

      if (!Utils.isValidGuid(args.options.runId)) {
        return `${args.options.runId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
      
    For text output mode, displays the name of the action, site script and the
    outcome of the action. For JSON output mode, displays all available
    information.

  Examples:
  
    List information about site scripts executed for the specified site design
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/team-a --runId b4411557-308b-4545-a3c4-55297d5cd8c8

  More information:

    SharePoint site design and site script overview
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview
`);
  }
}

module.exports = new SpoSiteDesignRunStatusGetCommand();