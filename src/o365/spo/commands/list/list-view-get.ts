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
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  viewId?: string;
  viewTitle?: string;
}

class SpoListViewGetCommand extends SpoCommand {

  public get name(): string {
    return commands.LIST_VIEW_GET;
  }

  public get description(): string {
    return 'Gets information about specific list view';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.listUrl = typeof args.options.listUrl !== 'undefined';
    telemetryProps.viewId = typeof args.options.viewId !== 'undefined';
    telemetryProps.viewTitle = typeof args.options.viewTitle !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    const baseRestUrl: string = `${args.options.webUrl}/_api/web`;
    let listRestUrl: string = '';

    if (args.options.listId) {

      listRestUrl = `/lists(guid'${encodeURIComponent(args.options.listId)}')`;

    } else if (args.options.listTitle) {

      listRestUrl = `/lists/getByTitle('${encodeURIComponent(args.options.listTitle as string)}')`;

    } else if (args.options.listUrl) {

      const listServerRelativeUrl: string = Utils.getServerRelativePath(args.options.webUrl, args.options.listUrl);

      listRestUrl = `/GetList('${encodeURIComponent(listServerRelativeUrl)}')`;
    }

    const viewRestUrl: string = `/views/${(args.options.viewId ? `getById('${encodeURIComponent(args.options.viewId)}')` : `getByTitle('${encodeURIComponent(args.options.viewTitle as string)}')`)}`;

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        const requestOptions: any = {
          url: `${baseRestUrl}${listRestUrl}${viewRestUrl}`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
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
      .then((result: Object): void => {

        cmd.log(result);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list is located'
      },
      {
        option: '--listId [listId]',
        description: 'ID of the list where the view is located. Specify only one of listTitle, listId or listUrl'
      },
      {
        option: '--listTitle [listTitle]',
        description: 'Title of the list where the view is located. Specify only one of listTitle, listId or listUrl'
      },
      {
        option: '--listUrl [listUrl]',
        description: 'Server or web relative url of the list where the view is located. Specify only one of listTitle, listId or listUrl'
      },
      {
        option: '--viewId [viewId]',
        description: 'ID of the view to get. Specify viewTitle or viewId but not both'
      },
      {
        option: '--viewTitle [viewTitle]',
        description: 'Title of the view to get. Specify viewTitle or viewId but not both'
      },
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

      if (!args.options.listId && !args.options.listTitle && !args.options.listUrl) {
        return `Specify listId, listTitle or listUrl`;
      }

      if (args.options.listId &&
        !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} in option listId is not a valid GUID`;
      }

      if (!args.options.viewId && !args.options.viewTitle) {
        return `Specify viewId or viewTitle`;
      }

      if (args.options.viewId && args.options.viewTitle) {
        return `Specify viewId or viewTitle but not both`;
      }

      if (args.options.viewId &&
        !Utils.isValidGuid(args.options.viewId)) {
        return `${args.options.viewId} in option viewId is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
    using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To get a list view, you have to first log in to SharePoint using
    the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
        
  Examples:

    Gets a list view located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} by its name
      ${chalk.grey(config.delimiter)} ${commands.LIST_VIEW_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --viewTitle 'All Items'

    Gets a list view located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} by its ID
      ${chalk.grey(config.delimiter)} ${commands.LIST_VIEW_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'Lists/My List' --viewId 330f29c5-5c4c-465f-9f4b-7903020ae1ce

    Gets a list view located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} by its name, but using ID with the listId option
      ${chalk.grey(config.delimiter)} ${commands.LIST_VIEW_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --listId 330f29c5-5c4c-465f-9f4b-7903020ae1c1 --viewTitle 'All Items'
   `);
  }
}

module.exports = new SpoListViewGetCommand();