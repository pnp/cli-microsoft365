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
import { ListInstanceCollection } from "./ListInstanceCollection";
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class ListListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_LIST;
  }

  public get description(): string {
    return 'Lists all available list in the specified site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Retrieving all lists in site at ${args.options.webUrl}...`);
        }

        let requestUrl: string;

        if (args.options.output === 'json') {
          requestUrl = `${args.options.webUrl}/_api/web/lists?$expand=RootFolder`;
        }
        else {
          requestUrl = `${args.options.webUrl}/_api/web/lists?$expand=RootFolder&$select=Title,Id,RootFolder/ServerRelativeURL`;
        }

        const requestOptions: any = {
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
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((listInstances: ListInstanceCollection): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(listInstances);
          cmd.log('');
        }

        if (args.options.output === 'json') {
          cmd.log(listInstances);
        }
        else {
          cmd.log(listInstances.value.map(l => {
            return {
              Title: l.Title,
              Url: l.RootFolder.ServerRelativeUrl,
              Id: l.Id
            };
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the lists to retrieve are located'
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
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
      using the ${chalk.blue(commands.CONNECT)} command.
  
  Remarks:
  
    To get all lists, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Return all lists located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_LIST} --webUrl https://contoso.sharepoint.com/sites/project-x
      `);
  }
}

module.exports = new ListListCommand();