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
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  label: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  syncToItems?: boolean;
  blockDelete?: boolean;
  blockEdit?: boolean;
}

class SpoListLabelSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_LABEL_SET;
  }

  public get description(): string {
    return 'Sets classification label on the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.listUrl = (!(!args.options.listUrl)).toString();
    telemetryProps.syncToItems = args.options.syncToItems || false;
    telemetryProps.blockDelete = args.options.blockDelete || false;
    telemetryProps.blockEdit = args.options.blockEdit || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    ((): Promise<string> => {
      let listRestUrl: string = '';

      if (args.options.listUrl) {
        const listServerRelativeUrl: string = Utils.getServerRelativePath(args.options.webUrl, args.options.listUrl);

        return Promise.resolve(listServerRelativeUrl);
      }
      else if (args.options.listId) {
        listRestUrl = `lists(guid'${encodeURIComponent(args.options.listId)}')/`;
      }
      else {
        listRestUrl = `lists/getByTitle('${encodeURIComponent(args.options.listTitle as string)}')/`;
      }

      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/${listRestUrl}?$expand=RootFolder&$select=RootFolder`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      };

      return request
        .get<ListInstance>(requestOptions)
        .then((listInstance: ListInstance): Promise<string> => {
          return Promise.resolve(listInstance.RootFolder.ServerRelativeUrl);
        });
    })()
      .then((listServerRelativeUrl: string): Promise<void> => {
        const listAbsoluteUrl: string = Utils.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);
        const requestUrl: string = `${args.options.webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`;
        const requestOptions: any = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          body: {
            listUrl: listAbsoluteUrl,
            complianceTagValue: args.options.label,
            blockDelete: args.options.blockDelete || false,
            blockEdit: args.options.blockEdit || false,
            syncToItems: args.options.syncToItems || false
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the list is located'
      },
      {
        option: '--label <label>',
        description: 'The label to set on the list'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'The title of the list on which to set the label. Specify only one of listTitle, listId or listUrl'
      },
      {
        option: '-l, --listId [listId]',
        description: 'The ID of the list on which to set the label. Specify only one of listTitle, listId or listUrl'
      },
      {
        option: '--listUrl [listUrl]',
        description: 'Server- or web-relative URL of the list on which to set the label. Specify only one of listTitle, listId or listUrl'
      },
      {
        option: '--syncToItems',
        description: 'Specify, to set the label on all items in the list'
      },
      {
        option: '--blockDelete',
        description: 'Specify, to disallow deleting items in the list'
      },
      {
        option: '--blockEdit',
        description: 'Specify, to disallow editing items in the list'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.listId && !args.options.listTitle && !args.options.listUrl) {
        return `Specify listId or listTitle or listUrl.`;
      }

      if (args.options.listId && !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} is not a valid GUID`;
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoListLabelSetCommand();