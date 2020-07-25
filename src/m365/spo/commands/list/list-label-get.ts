import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ListInstance } from './ListInstance';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
}

class SpoListLabelGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_LABEL_GET;
  }

  public get description(): string {
    return 'Gets label set on the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      const list: string = args.options.listId ? encodeURIComponent(args.options.listId as string) : encodeURIComponent(args.options.listTitle as string);
      cmd.log(`Getting label set on the list ${list} in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.listId) {
      if (this.debug) {
        cmd.log(`Retrieving List Url from Id '${args.options.listId}'...`);
      }

      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')?$expand=RootFolder&$select=RootFolder`;
    }
    else {
      if (this.debug) {
        cmd.log(`Retrieving List Url from Title '${args.options.listTitle}'...`);
      }

      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')?$expand=RootFolder&$select=RootFolder`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<ListInstance>(requestOptions)
      .then((listInstance: ListInstance): Promise<any> => {
        const listAbsoluteUrl: string = Utils.getAbsoluteUrl(args.options.webUrl, listInstance.RootFolder.ServerRelativeUrl);
        const requestUrl: string = `${args.options.webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`;
        const requestOptions: any = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata'
          },
          json: true,
          body: {
            listUrl: listAbsoluteUrl
          }
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (res['odata.null'] !== true) {
          cmd.log(res);
        }

        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list to get the label from is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list to get the label from. Specify either listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list to get the label from. Specify either listId or listTitle but not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
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
}

module.exports = new SpoListLabelGetCommand();