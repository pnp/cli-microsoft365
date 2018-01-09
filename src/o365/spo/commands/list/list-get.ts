import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { ContextInfo } from '../../spo';
import Utils from '../../../../Utils';
import { ListInstance } from "./ListInstance";
import Auth from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  title?: string;
}

class ListGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_GET;
  }

  public get description(): string {
    return 'Gets information about the specific list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.title = (!(!args.options.title)).toString();
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
    .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
    .then((accessToken: string): Promise<ContextInfo> => {
      siteAccessToken = accessToken;

      if (this.debug) {
        cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
      }

      return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
    })
    .then((res: ContextInfo): Promise<ListInstance> => {
      if (this.debug) {
        cmd.log('Response:');
        cmd.log(res);
        cmd.log('');
      }

      if (this.verbose) {
        cmd.log(`Retrieving information for list in site at ${args.options.webUrl}...`);
      }

      let requestOptions: any = {};
      let requestUrl: string = '';

      if (args.options.id) {
        requestUrl = `${auth.site.url}/_api/web/lists(guid'${args.options.id}')`;
      }
      else {
        requestUrl = `${auth.site.url}/_api/web/lists/GetByTitle('${args.options.title}')`;
      }

      requestOptions = {
        url: requestUrl,
        method: 'GET',
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${siteAccessToken}`,
          'accept': 'application/json;odata=nometadata'
        }),
        json: true
      };

      if (this.debug) {
        cmd.log('Executing web request...');
        cmd.log(JSON.stringify(requestOptions));
        cmd.log('');
      }

      return request.get(requestOptions);
    })
    .then((listInstance: ListInstance): void => {
      if (this.debug) {
        cmd.log('Response:');
        cmd.log(JSON.stringify(listInstance));
        cmd.log('');
      }

      cmd.log(listInstance);
      cb();
    }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list to retrieve is located'
      },
      {
        option: '-i, --id <id>',
        description: 'List id. Specify either id or title but not both'
      },
      {
        option: '-t, --title <title>',
        description: 'List title. Specify either id or title but not both'
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

      if (args.options.id) {
        if (!Utils.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }
      }

      if (args.options.id && args.options.title) {
        return 'Specify id or title, but not both';
      }

      if (!args.options.id && !args.options.title) {
        return 'Specify id or title, one is required';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
        using the ${chalk.blue(commands.CONNECT)} command.
        
        Examples:
        
          Return information about a list by id
            ${chalk.grey(config.delimiter)} ${commands.LIST_GET} -u https://contoso.sharepoint.com/sites/project-x -i 0CD891EF-AFCE-4E55-B836-FCE03286CCCF

          Return information about a list by title
            ${chalk.grey(config.delimiter)} ${commands.LIST_GET} -u https://contoso.sharepoint.com/sites/project-x -t Documents
      `);
    }
}

module.exports = new ListGetCommand();