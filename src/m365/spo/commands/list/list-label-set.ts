import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListInstance } from './ListInstance';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    ((): Promise<string> => {
      let listRestUrl: string = '';

      if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);

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
        responseType: 'json'
      };

      return request
        .get<ListInstance>(requestOptions)
        .then((listInstance: ListInstance): Promise<string> => {
          return Promise.resolve(listInstance.RootFolder.ServerRelativeUrl);
        });
    })()
      .then((listServerRelativeUrl: string): Promise<void> => {
        const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);
        const requestUrl: string = `${args.options.webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`;
        const requestOptions: any = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          data: {
            listUrl: listAbsoluteUrl,
            complianceTagValue: args.options.label,
            blockDelete: args.options.blockDelete || false,
            blockEdit: args.options.blockEdit || false,
            syncToItems: args.options.syncToItems || false
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--label <label>'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--syncToItems'
      },
      {
        option: '--blockDelete'
      },
      {
        option: '--blockEdit'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.listId && !args.options.listTitle && !args.options.listUrl) {
      return `Specify listId or listTitle or listUrl.`;
    }

    if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
      return `${args.options.listId} is not a valid GUID`;
    }

    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoListLabelSetCommand();