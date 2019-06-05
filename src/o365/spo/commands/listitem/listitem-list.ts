import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate,
  CommandTypes
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ListItemInstanceCollection } from './ListItemInstanceCollection';
import { ContextInfo } from '../../spo';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  fields?: string;
  filter?: string;
  pageNumber?: string;
  pageSize?: string;
  query?: string;
  title?: string;
  webUrl: string;
}

class SpoListItemListCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_LIST;
  }

  public get description(): string {
    return 'Gets a list of items from the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.fields = typeof args.options.fields !== 'undefined';
    telemetryProps.filter = typeof args.options.filter !== 'undefined';
    telemetryProps.pageNumber = typeof args.options.pageNumber !== 'undefined';
    telemetryProps.pageSize = typeof args.options.pageSize !== 'undefined';
    telemetryProps.query = typeof args.options.query !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const listIdArgument = args.options.id || '';
    const listTitleArgument = args.options.title || '';

    let formDigestValue: string = '';

    const fieldsArray: string[] = args.options.fields ? args.options.fields.split(",")
      : (!args.options.output || args.options.output === "text") ? ["Title", "Id"] : []

    const listRestUrl: string = (args.options.id ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);

    ((): Promise<any> => {
      if (args.options.query) {
        if (this.debug) {
          cmd.log(`getting request digest for query request`);
        }

        return this.getRequestDigest(args.options.webUrl);
      }
      else {
        return Promise.resolve();
      }
    })()
      .then((res: ContextInfo): Promise<any> => {
        formDigestValue = args.options.query ? res.FormDigestValue : '';

        if (args.options.pageNumber && Number(args.options.pageNumber) > 0) {
          const rowLimit: string = `$top=${Number(args.options.pageSize) * Number(args.options.pageNumber)}`;
          const filter: string = args.options.filter ? `$filter=${encodeURIComponent(args.options.filter)}` : ``;
          const fieldSelect: string = `?$select=Id&${rowLimit}&${filter}`;

          const requestOptions: any = {
            url: `${listRestUrl}/items${fieldSelect}`,
            headers: {
              'accept': 'application/json;odata=nometadata',
              'X-RequestDigest': formDigestValue
            },
            json: true
          };

          return request.get(requestOptions);
        }
        else {
          return Promise.resolve();
        }
      })
      .then((res: any): Promise<ListItemInstanceCollection> => {
        const skipTokenId = (res && res.value && res.value.length && res.value[res.value.length - 1]) ? res.value[res.value.length - 1].Id : 0
        const skipToken: string = (args.options.pageNumber && Number(args.options.pageNumber) > 0 && skipTokenId > 0) ? `$skiptoken=Paged=TRUE%26p_ID=${res.value[res.value.length - 1].Id}` : ``;
        const rowLimit: string = args.options.pageSize ? `$top=${args.options.pageSize}` : ``
        const filter: string = args.options.filter ? `$filter=${encodeURIComponent(args.options.filter)}` : ``
        const fieldSelect: string = fieldsArray.length > 0 ?
          `?$select=${encodeURIComponent(fieldsArray.join(","))}&${rowLimit}&${skipToken}&${filter}` :
          `?${rowLimit}&${skipToken}&${filter}`
        const requestBody: any = args.options.query ?
          {
            "query": {
              "ViewXml": args.options.query
            }
          }
          : ``;

        const requestOptions: any = {
          url: `${listRestUrl}/${args.options.query ? `GetItems` : `items${fieldSelect}`}`,
          headers: {
            'accept': 'application/json;odata=nometadata',
            'X-RequestDigest': formDigestValue
          },
          json: true,
          body: requestBody
        };

        return args.options.query ? request.post(requestOptions) : request.get(requestOptions);
      })
      .then((listItemInstances: ListItemInstanceCollection): void => {
        if (args.options.output === 'json') {
          cmd.log(listItemInstances.value);
        }
        else {
          cmd.log(listItemInstances.value.map(l => {
            if ((<any>l)["ID"] && l["Id"]) delete (<any>l)["ID"];
            return l;
          }));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
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
        option: '-n, --pageNumber [pageNumber]',
        description: 'Page number to return if pageSize is specified (first page is indexed as value of 0)'
      },
      {
        option: '-q, --query [query]',
        description: 'CAML query to use to retrieve items. Will ignore pageSize and pageNumber if specified'
      },
      {
        option: '-f, --fields [fields]',
        description: 'Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested. Specify query or fields but not both'
      },
      {
        option: '-l, --filter [odataFilter]',
        description: 'OData filter to use to query the list of items with. Specify query or filter but not both'
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
        'pageNumber',
        'fields',
        'filter',
      ],
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
        return `Specify list id or title`;
      }

      if (args.options.id && args.options.title) {
        return `Specify list id or title but not both`;
      }

      if (args.options.query && args.options.fields) {
        return `Specify query or fields but not both`;
      }

      if (args.options.query && args.options.pageSize) {
        return `Specify query or pageSize but not both`;
      }

      if (args.options.query && args.options.pageNumber) {
        return `Specify query or pageNumber but not both`;
      }

      if (args.options.pageSize && isNaN(Number(args.options.pageSize))) {
        return `pageSize must be numeric`;
      }

      if (args.options.pageNumber && !args.options.pageSize) {
        return `pageSize must be specified if pageNumber is specified`;
      }

      if (args.options.pageNumber && isNaN(Number(args.options.pageNumber))) {
        return `pageNumber must be numeric`;
      }

      if (args.options.id &&
        !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} in option id is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
  
    ${chalk.grey('pageNumber')} is specified as a 0-based index. A value of ${chalk.grey('2')} returns the third
    page of items.
        
  Examples:
  
    Get all items from a list named ${chalk.grey('Demo List')}
      ${commands.LISTITEM_LIST} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x

    From a list named ${chalk.grey('Demo List')} get all items with title ${chalk.grey('Demo list item')}
    using a CAML query
      ${commands.LISTITEM_LIST} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --query "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo list item</Value></Eq></Where></Query></View>"
    
    Get all items from a list with ID ${chalk.grey('935c13a0-cc53-4103-8b48-c1d0828eaa7f')} 
      ${commands.LISTITEM_LIST} --id 935c13a0-cc53-4103-8b48-c1d0828eaa7f --webUrl https://contoso.sharepoint.com/sites/project-x

    Get all items from list named ${chalk.grey('Demo List')}. For each item, retrieve the value
    of the ${chalk.grey('ID')}, ${chalk.grey('Title')} and ${chalk.grey('Modified')} fields
      ${commands.LISTITEM_LIST} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --fields "ID,Title,Modified"

    From a list named ${chalk.grey('Demo List')} get all items with title ${chalk.grey('Demo list item')}
    using an OData filter 
      ${commands.LISTITEM_LIST} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --filter "Title eq 'Demo list item'"

    From a list named ${chalk.grey('Demo List')} get the third batch of 10 items
      ${commands.LISTITEM_LIST} --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --pageSize 10 --pageNumber 2
   `);
  }
}

module.exports = new SpoListItemListCommand();