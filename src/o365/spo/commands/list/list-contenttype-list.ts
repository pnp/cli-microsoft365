import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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
}

class SpoListContentTypeListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_CONTENTTYPE_LIST;
  }

  public get description(): string {
    return 'Lists content types configured on the list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): Promise<any> => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving information about content types...`);
        }

        if (this.verbose) {
          const list: string = args.options.listId ? encodeURIComponent(args.options.listId as string) : encodeURIComponent(args.options.listTitle as string);
          cmd.log(`Retrieving content types information for list ${list} in site at ${args.options.webUrl}...`);
        }

        let requestUrl: string = '';

        if (args.options.listId) {
          requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')/ContentTypes`;
        }
        else {
          requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')/ContentTypes`;
        }

        const requestOptions: any = {
          url: requestUrl,
          method: 'GET',
          headers: {
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res: any): void => {
        if (args.options.output === 'json') {
          cmd.log(res.value);
        }
        else {
          cmd.log(res.value.map((ct: any) => {
            return {
              StringId: ct.StringId,
              Name: ct.Name,
              Hidden: ct.Hidden,
              ReadOnly: ct.ReadOnly,
              Sealed: ct.Sealed
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
        description: 'URL of the site where the list is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list for which to list configured content types. Specify listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list for which to list configured content types. Specify listId or listTitle but not both'
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

      if (args.options.listId) {
        if (!Utils.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }
      }

      if (args.options.listId && args.options.listTitle) {
        return 'Specify listId or listTitle, but not both';
      }

      if (!args.options.listId && !args.options.listTitle) {
        return 'Specify listId or listTitle, one is required';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to SharePoint,
    using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To list all content types configured on a list, you have to first log in
    to SharePoint using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    List all content types configured on the list with ID
    ${chalk.grey('0cd891ef-afce-4e55-b836-fce03286cccf')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_CONTENTTYPE_LIST} --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf

    List all content types configured on the list with title ${chalk.grey('Documents')} located
    in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_CONTENTTYPE_LIST} --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents
      `);
  }
}

module.exports = new SpoListContentTypeListCommand();