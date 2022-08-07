import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  listItemId: string;
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --listItemId <listItemId>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const id: number = parseInt(args.options.listItemId);
        if (isNaN(id)) {
          return `${args.options.listItemId} is not a valid list item ID`;
        }

        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['listId', 'listTitle']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const listIdArgument: string = args.options.listId || '';
    const listTitleArgument: string = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitleArgument)}')`);

    let formDigestValue: string = '';
    let environmentListId: string = '';

    try {
      if (typeof args.options.listId !== 'undefined') {
        environmentListId = args.options.listId;
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`Getting list id...`);
        }
        const listRequestOptions: any = {
          url: `${listRestUrl}/id`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
  
        const idResp = await request.get<{ value: string; }>(listRequestOptions);
        environmentListId = idResp.value;
      }

      if (this.debug) {
        logger.logToStderr(`getting request digest for request`);
      }

      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      formDigestValue = reqDigest.FormDigestValue;

      const objectIdentity = await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);

      if (this.verbose) {
        logger.logToStderr(`Undeclare list item as a record in list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}...`);
      }

      const requestOptions: any = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'Content-Type': 'text/xml',
          'X-RequestDigest': formDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="UndeclareItemAsRecord" Id="53"><Parameters><Parameter ObjectPathId="49" /></Parameters></StaticMethod></Actions><ObjectPaths><Identity Id="49" Name="${objectIdentity.objectIdentity}:list:${environmentListId}:item:${args.options.listItemId},1" /></ObjectPaths></Request>`
      };

      await request.post(requestOptions);
      // REST post call doesn't return anything
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoListItemRecordUndeclareCommand();