import auth from '../../SpoAuth';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  asTask: boolean;
  id: string;
  webUrl: string;
}

class SpoSiteDesignApplyCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_APPLY}`;
  }

  public get description(): string {
    return 'Applies a site design to an existing site collection';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.asTask = args.options.asTask || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): Promise<any> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Applying site design...`);
        }

        const requestBody: any = {
          siteDesignId: args.options.id,
          webUrl: args.options.webUrl
        };

        const requestOptions: any = {
          url: `${auth.site.url}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.${args.options.asTask ? 'AddSiteDesignTask' : 'ApplySiteDesign'}`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          body: requestBody,
          json: true
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (res.value) {
          cmd.log(res.value);
        }
        else {
          cmd.log(res);
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
        option: '-i, --id <id>',
        description: 'The ID of the site design to apply'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site to apply the site design to'
      },
      {
        option: '--asTask',
        description: 'Apply site design as task. Required for large site designs'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }
      else {
        const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
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

    To apply a site design to an existing site collection, you have to first
    log in to a SharePoint site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

  Examples:

    Apply the site design with ID ${chalk.grey('9b142c22-037f-4a7f-9017-e9d8c0e34b98')}
    to the site collection ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${this.name} --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --webUrl https://contoso.sharepoint.com/sites/project-x

    Apply large site design to the specified site
      ${chalk.grey(config.delimiter)} ${this.name} --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --webUrl https://contoso.sharepoint.com/sites/project-x --asTask

  More information:

    SharePoint site design and site script overview
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview
`);
  }
}

module.exports = new SpoSiteDesignApplyCommand();