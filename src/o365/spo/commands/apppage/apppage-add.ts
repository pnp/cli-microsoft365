import auth from '../../SpoAuth';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  title: string;
  webPartData: string;
  addToQuickLaunch: boolean;
}

class SpoAppPageAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.APPPAGE_ADD}`;
  }

  public get description(): string {
    return 'Creates a single-part app page';
  }
  //todo:remove
  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }
  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): Promise<ContextInfo> => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): Promise<{}> => {
        const requestOptions: any = {
          url: `${auth.site.url}/_api/sitepages/Pages/CreateFullPageApp`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          json:true,
          body: {
          title:args.options.title,
          addToQuickLaunch:args.options.addToQuickLaunch?true:false,
          webPartDataAsJson:args.options.webPartData
        }
      };
        return request.post(requestOptions);
      })
      .then((res: any): void => {
        cmd.log(res);

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
        description: 'The URL of the site where the page should be created'
      },
      {
        option: '-t, --title <title>',
        description: 'The title of the page to create'
      },
      {
        option: '-d, --webPartData <webPartData>',
        description: 'JSON string of the web part to put on the page'
      },
      {
        option: '--addToQuickLaunch',
        description: 'Set, to add the page to the quick launch'
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
      if (!args.options.title) {
        return 'Required parameter title missing';
      }
      if (!args.options.webPartData) {
        return 'Required parameter webPartData missing';
      }
      try {
       JSON.parse(args.options.webPartData);
      }
      catch (e) {
        return `The webPartData passed is not a valid JSON string`;
      }
      return true;
    };
  }
  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site using the
      ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To a single-part app page, you have to first log in to a SharePoint site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
    If you want to add the single-part app page to quicklaunch, use the ${chalk.blue('addToQuickLaunch')}
    flag.

  Examples:
  
    Create a single-part app page in a web with url https://contoso.sharepoint.com, webpart data are stored in the ${chalk.grey('$webPartData')} variable
      ${chalk.grey(config.delimiter)} ${this.name} --title "app page" --webUrl "https://contoso.sharepoint.com" --webPartData $webPartData --addToQuickLaunch 
`);
  }
}

module.exports = new SpoAppPageAddCommand();