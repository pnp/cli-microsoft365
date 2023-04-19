import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListItemInstance } from './ListItemInstance';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  id?: string;
  uniqueId?: string;
  properties?: string;
  withPermissions?: boolean;
}

class SpoListItemGetCommand extends SpoCommand {
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LISTITEM_GET;
  }

  public get description(): string {
    return 'Gets a list item from the specified list';
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
        id: typeof args.options.id !== 'undefined',
        uniqueId: typeof args.options.uniqueId !== 'undefined',
        withPermissions: !!args.options.withPermissions
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--uniqueId [uniqueId]'
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
        option: '-p, --properties [properties]'
      },
      {
        option: '--withPermissions'
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

        if (args.options.id &&
          isNaN(parseInt(args.options.id))) {
          return `${args.options.id} is not a number`;
        }

        if (args.options.uniqueId &&
          !validation.isValidGuid(args.options.uniqueId)) {
          return `${args.options.uniqueId} in option uniqueId is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push(
      'webUrl',
      'listId',
      'listTitle',
      'id',
      'uniqueId',
      'properties'
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['listId', 'listTitle', 'listUrl'] },
      { options: ['id', 'uniqueId'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let requestUrl = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    const propertiesSelect: string[] = args.options.properties ? args.options.properties.split(',') : [];
    const propertiesWithSlash: string[] = propertiesSelect.filter(item => item.includes('/'));
    const propertiesToExpand: string[] = propertiesWithSlash.map(e => e.split('/')[0]);
    const expandPropertiesArray: string[] = propertiesToExpand.filter((item, pos) => propertiesToExpand.indexOf(item) === pos);
    const fieldExpand: string = expandPropertiesArray.length > 0 ? `&$expand=${expandPropertiesArray.join(",")}` : ``;

    if (args.options.id) {
      requestUrl += `/items(${args.options.id})`;
    }
    else {
      requestUrl += `/GetItemByUniqueId(guid'${args.options.uniqueId}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$select=${formatting.encodeQueryParameter(propertiesSelect.join(","))}${fieldExpand}`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const itemProperties = await request.get<any>(requestOptions);
      if (args.options.withPermissions) {
        requestOptions.url = `${requestUrl}/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
        const response = await request.get<{ value: any[] }>(requestOptions);
        response.value.forEach((r: any) => {
          r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
        });
        itemProperties.RoleAssignments = response.value;
      }
      delete itemProperties['ID'];
      logger.log(<ListItemInstance>itemProperties);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListItemGetCommand();
