import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
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

export interface ListItemIsRecord {
  CallInfo: any;
  CallObjectId: any;
  IsRecord: any;
}

class SpoListItemIsRecordCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ISRECORD;
  }

  public get description(): string {
    return 'Checks if the specified list item is a record';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const clientSvcCommons: ClientSvc = new ClientSvc(cmd, this.debug);
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    const listIdArgument: string = args.options.listId || '';
    const listTitleArgument: string = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);

    let formDigestValue: string = '';
    let listId: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    ((): Promise<{ Id: string; }> => {
      if (typeof args.options.listId !== 'undefined') {
        if (this.verbose) {
          cmd.log(`List Id passed in as an argument.`);
        }

        return Promise.resolve({ Id: args.options.listId });
      }

      if (this.verbose) {
        cmd.log(`Getting list id...`);
      }
      const requestOptions: any = {
        url: `${listRestUrl}?$select=Id`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        json: true
      }

      return request.get(requestOptions);
    })()
      .then((res: { Id: string }): Promise<ContextInfo> => {
        listId = res.Id;

        if (this.debug) {
          cmd.log(`Getting request digest for request`);
        }

        return this.getRequestDigest(args.options.webUrl);
      })
      .then((res: ContextInfo): Promise<IdentityResponse> => {
        formDigestValue = res.FormDigestValue;
        return clientSvcCommons.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
      })
      .then((webIdentityResp: IdentityResponse): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Checking if list item is a record in list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}...`);
        }

        const requestBody = this.getIsRecordRequestBody(webIdentityResp.objectIdentity, listId, args.options.id)
        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'Content-Type': 'text/xml',
            'X-RequestDigest': formDigestValue,
          },
          body: requestBody
        };

        return request.post<string>(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];

        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
        }
        else {
          const result: boolean = json[json.length - 1];
          cmd.log(result);
          cb();
        }
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  };

  private getIsRecordRequestBody(webIdentity: string, listId: string, id: string): string {
    const requestBody: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
            <Actions>
              <StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="IsRecord" Id="1"><Parameters><Parameter ObjectPathId="14" /></Parameters></StaticMethod>
            </Actions>
            <ObjectPaths>
              <Identity Id="14" Name="${webIdentity}:list:${listId}:item:${id},1" />
            </ObjectPaths>
          </Request>`;

    return requestBody;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the list is located'
      },
      {
        option: '-i, --id <id>',
        description: 'The ID of the list item to check if it is a record'
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

      if (id < 1) {
        return `Item ID must be a positive number`;
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

module.exports = new SpoListItemIsRecordCommand();