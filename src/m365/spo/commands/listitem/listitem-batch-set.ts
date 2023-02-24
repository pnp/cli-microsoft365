import * as fs from 'fs';
import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListInstance } from '../list/ListInstance';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filePath: string;
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  systemUpdate?: boolean;
}

class SpoListItemBatchSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_BATCH_SET;
  }

  public get description(): string {
    return 'Updates list items in a batch';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        systemUpdate: !!args.options.systemUpdate
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-p, --filePath <filePath>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--systemUpdate'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        if (!fs.existsSync(args.options.filePath)) {
          return `File with path ${args.options.filePath} does not exist`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push(
      'webUrl',
      'filePath',
      'listId',
      'listTitle',
      'listUrl'
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Starting to create batch items from csv at path ${args.options.filePath}`);
      }
      const csvContent = fs.readFileSync(args.options.filePath, 'utf8');
      const jsonContent: any[] = formatting.parseCsvToJson(csvContent);
      // check if ID Column exists
      if ('ID' in jsonContent[0] === false && 'id' in jsonContent[0] === false && 'Id' in jsonContent[0] === false) {
        throw 'Please make sure that the CSV has a column with the name ID';
      }

      const idColumn = 'ID' in jsonContent[0] ? "ID" : 'Id' in jsonContent[0] ? "Id" : "id";

      const formDigestValue = (await spo.getRequestDigest(args.options.webUrl)).FormDigestValue;
      const objectIdentity = (await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue)).objectIdentity;
      const listId = await this.getListId(args.options);

      let objectPaths = [];
      let actions = [];
      let index = 1;

      for await (const [batchIndex, row] of jsonContent.entries()) {
        objectPaths.push(`<Identity Id="${index}" Name="${objectIdentity}:list:${listId}:item:${row[idColumn]},1" />`);

        const [actionString, updatedIndex] = this.mapActions(index, row, args.options.systemUpdate);
        index = updatedIndex;
        actions.push(actionString);

        if (objectPaths.length === 50) {
          if (this.verbose) {
            logger.logToStderr(`Writing away batch of items, currently at: ${batchIndex + 1}/${jsonContent.length}.`);
          }

          await this.sendBatchRequest(args.options.webUrl, this.getRequestBody(objectPaths, actions));
          objectPaths = [];
          actions = [];
        }
      }

      if (objectPaths.length) {
        if (this.verbose) {
          logger.logToStderr(`Writing away ${objectPaths.length} items.`);
        }

        await this.sendBatchRequest(args.options.webUrl, this.getRequestBody(objectPaths, actions));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRequestBody(objectPaths: string[], actions: string[]): string {
    return `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${actions.join('')}</Actions><ObjectPaths>${objectPaths.join('')}</ObjectPaths></Request>`;
  }

  private async sendBatchRequest(webUrl: string, requestBody: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'Content-Type': 'text/xml'
      },
      data: requestBody
    };

    await request.post(requestOptions);
  }

  private mapActions(index: number, row: any, systemUpdate?: boolean): [string, number] {
    const objectPathId = index;
    let actionString = '';
    const excludeOptions = ['ID', 'id', 'Id'];

    Object.keys(row).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        actionString += `<Method Name="ParseAndSetFieldValue" Id="${index++}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${key}</Parameter><Parameter Type="String">${(<any>row)[key].toString()}</Parameter></Parameters></Method>`;
      }
    });

    actionString += `<Method Name="${systemUpdate ? 'System' : ''}Update" Id="${index++}" ObjectPathId="${objectPathId}"/>`;
    return [actionString, index];
  }

  private async getListId(options: Options): Promise<string> {
    let listUrl = `${options.webUrl}/_api/web`;
    if (options.listId) {
      return options.listId;
    }
    else if (options.listTitle) {
      listUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }
    else if (options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      listUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }
    listUrl += '?$select=Id';

    const requestOptions: CliRequestOptions = {
      url: listUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const listInstance = await request.get<ListInstance>(requestOptions);
    return listInstance.Id;
  }
}

module.exports = new SpoListItemBatchSetCommand();