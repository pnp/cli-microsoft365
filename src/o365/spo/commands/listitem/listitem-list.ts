import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate,
  CommandTypes
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';
import { ListItemInstance } from './ListItemInstance';
import { ContextInfo } from '../../spo';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  title?: string;
  pageSize?: string;
  query?: string;
  filter?: string;
  fields?: string;
}

class SpoListItemGetCommand extends SpoCommand {
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LISTITEM_LIST;
  }

  public get description(): string {
    return 'Get list items from the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    const listIdArgument = args.options.id || '';
    const listTitleArgument = args.options.title || '';
    
    let siteAccessToken: string = '';
    let formDigestValue: string = '';

    const listRestUrl: string = (args.options.id ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise | Promise<any> => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
          cmd.log(``);
          cmd.log(`auth object:`);
          cmd.log(auth);
        }

        if (args.options.query) {
          if (this.debug) {
            cmd.log(`getting request digest for query request`);
          }

          return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
        }
        else {
          return Promise.resolve();
        }
      })
      .then((res: ContextInfo): request.RequestPromise<any> => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }
        
        formDigestValue = args.options.query ? res['FormDigestValue'] : '';
        const fieldSelect: string = args.options.fields ?
          `?$select=${encodeURIComponent(args.options.fields)}` :
          (
            (!args.options.output || args.options.output === 'text') ?
              `?$select=Id,Title` :
              ``
          )

        const requestBody: any = args.options.query ?
            {
              "query": { 
                "__metadata": { 
                  "type": "SP.CamlQuery"
                }, 
                "ViewXml": args.options.query 
              } 
            }
          : ``;
        
        const requestOptions: any = {
          url: `${listRestUrl}/${args.options.query ? `GetItems` : `items/${fieldSelect}`}`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata',
            'X-RequestDigest': args.options.query ? formDigestValue : ''
          }),
          json: true,
          body: requestBody
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return args.options.query ? request.get(requestOptions) : request.post(requestOptions);
      })
      .then((response: any): void => {
        (!args.options.output || args.options.output === 'text') && delete response["ID"];
        cmd.log(<ListItemInstance>response);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --webUrl <webUrl>',
        description: 'URL of the site where the list from which to retrieve items is located'
      },
      {
        option: '-i, --id [listId]',
        description: 'ID of the list from which to retrieve items. Specify id or title but not both'
      },
      {
        option: '-t, --title [listTitle]',
        description: 'Title of the list from which to retrieve items. Specify id or title but not both'
      },
      {
        option: '-s, --pageSize [pageSize]',
        description: 'The number of items to retrieve per page request'
      },
      {
        option: '-q, --query [query]',
        description: 'The CAML query to use to retrieve items. Will ignore pageSize if specified'
      },
      {
        option: '-f, --fields [fields]',
        description: 'Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public types(): CommandTypes {
    return {
      string: [
        'webUrl',
        'id',
        'title',
        'query',
        'pageSize',
        'fields',
      ]
    };
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

      if (!args.options.id && !args.options.title) {
        return `Specify listId or listTitle`;
      }

      if (args.options.id && args.options.title) {
        return `Specify listId or listTitle but not both`;
      }

      if (args.options.id &&
        !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} in option listId is not a valid GUID`;
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
  
    To get an items from a list, you have to first connect to SharePoint using
    the ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Get the items in list with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_GET} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x

   `);
  }

}

module.exports = new SpoListItemGetCommand();