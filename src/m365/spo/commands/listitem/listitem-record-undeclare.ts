import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import { ClientSvc, IdentityResponse } from '../../ClientSvc';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  listId?: string;
  listTitle?: string;
  webUrl: string;
}

class SpoListItemRecordUndeclareCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_RECORD_UNDECLARE;
  }

  public get description(): string {
    return 'Undeclares list item as a record';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const clientSvcCommons: ClientSvc = new ClientSvc(cmd, this.debug);
    const listIdArgument: string = args.options.listId || '';
    const listTitleArgument: string = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);

    let formDigestValue: string = '';
    let environmentListId: string = '';

    ((): Promise<{ value: string; }> => {
      if (typeof args.options.listId !== 'undefined') {
        return Promise.resolve({ value: args.options.listId });
      }

      if (this.verbose) {
        cmd.log(`Getting list id...`);
      }
      const listRequestOptions: any = {
        url: `${listRestUrl}/id`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      };

      return request.get(listRequestOptions);
    })()
      .then((res: { value: string }): Promise<ContextInfo> => {
        environmentListId = res.value;

        if (this.debug) {
          cmd.log(`getting request digest for request`);
        }

        return this.getRequestDigest(args.options.webUrl);
      })
      .then((res: ContextInfo): Promise<IdentityResponse> => {
        formDigestValue = res.FormDigestValue;

        return clientSvcCommons.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
      })
      .then((objectIdentity: IdentityResponse): Promise<void> => {
        if (this.verbose) {
          cmd.log(`Undeclare list item as a record in list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'Content-Type': 'text/xml',
            'X-RequestDigest': formDigestValue,
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="UndeclareItemAsRecord" Id="53"><Parameters><Parameter ObjectPathId="49" /></Parameters></StaticMethod></Actions><ObjectPaths><Identity Id="49" Name="${objectIdentity.objectIdentity}:list:${environmentListId}:item:${args.options.id},1" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        // REST post call doesn't return anything
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the list is located'
      },
      {
        option: '-i, --id <id>',
        description: 'ID of the list item to be undeclared as a record.'
      },
      {
        option: '-l, --listId [listId]',
        description: 'The ID of the list where the item is located. Specify listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'The title of the list where the item is located. Specify listId or listTitle but not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const id: number = parseInt(args.options.id);
      if (isNaN(id)) {
        return `${args.options.id} is not a valid list item ID`;
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.listId && !args.options.listTitle) {
        return `Specify listId or listTitle`;
      }

      if (args.options.listId && args.options.listTitle) {
        return `Specify listId or listTitle but not both`;
      }

      if (args.options.listId &&
        !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} in option listId is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new SpoListItemRecordUndeclareCommand();