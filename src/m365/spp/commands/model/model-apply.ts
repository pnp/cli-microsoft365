import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spp, SppModel } from '../../../../utils/spp.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { ListInstance } from '../../../spo/commands/list/ListInstance.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  contentCenterUrl: string;
  siteUrl: string;
  id?: string;
  title?: string;
  listTitle?: string;
  listId?: string;
  listUrl?: string;
  viewOption?: string;
}

class SppModelApplyCommand extends SpoCommand {
  public readonly viewOptions: string[] = ['NewViewAsDefault', 'DoNotChangeDefault', 'TileViewAsDefault'];

  public get name(): string {
    return commands.MODEL_APPLY;
  }

  public get description(): string {
    return 'Applies (or syncs) a trained document understanding model to a document library';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        viewOption: typeof args.options.viewOption !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '--contentCenterUrl <contentCenterUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--viewOption [viewOption]',
        autocomplete: this.viewOptions
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} in parameter id is not a valid GUID`;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in parameter listId is not a valid GUID`;
        }

        if (typeof args.options.viewOption !== 'undefined') {
          if (!this.viewOptions.some(viewOption => viewOption.toLocaleLowerCase() === args.options.viewOption?.toLowerCase())) {
            return `The value of parameter viewOption must be ${this.viewOptions.join(', ')}`;
          }
        }

        const isContentCenterUrlValid = validation.isValidSharePointUrl(args.options.contentCenterUrl);

        if (isContentCenterUrlValid !== true) {
          return isContentCenterUrlValid;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title'] });
    this.optionSets.push({ options: ['listTitle', 'listId', 'listUrl'] });
  }

  #initTypes(): void {
    this.types.string.push('contentCenterUrl', 'siteUrl', 'id', 'title', 'listTitle', 'listId', 'listUrl', 'viewOption');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.log(`Applying a model to a document library...`);
      }

      const contentCenterUrl = urlUtil.removeTrailingSlashes(args.options.contentCenterUrl);
      await spp.assertSiteIsContentCenter(contentCenterUrl);

      const model = await this.getModel(contentCenterUrl, args);
      const listInstance = await this.getDocumentLibraryInfo(args);

      if (listInstance.BaseType !== 1) {
        throw Error(`The specified list is not a document library.`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${contentCenterUrl}/_api/machinelearning/publications`,
        headers: {
          accept: 'application/json;odata=nometadata',
          "Content-Type": 'application/json;odata=verbose'
        },
        responseType: 'json',
        data: {
          __metadata: { type: 'Microsoft.Office.Server.ContentCenter.SPMachineLearningPublicationsEntityData' },
          Publications: {
            results: [
              {
                ModelUniqueId: model.UniqueId,
                TargetSiteUrl: args.options.siteUrl,
                TargetWebServerRelativeUrl: urlUtil.getServerRelativePath(args.options.siteUrl, ''),
                TargetLibraryServerRelativeUrl: listInstance.RootFolder.ServerRelativeUrl,
                ViewOption: args.options.viewOption ? args.options.viewOption : "NewViewAsDefault"
              }
            ]
          }
        }
      };

      const result = await request.post(requestOptions);
      const resutlDetails = (result as any)?.Details;

      if (resutlDetails?.length > 0 && resutlDetails[0]?.ErrorMessage) {
        throw (result as any).Details[0].ErrorMessage;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getModel(contentCenterUrl: string, args: CommandArgs): Promise<SppModel> {
    let requestUrl = `${contentCenterUrl}/_api/machinelearning/models/`;

    if (args.options.title) {
      let requestTitle = args.options.title.toLowerCase();

      if (!requestTitle.endsWith('.classifier')) {
        requestTitle += '.classifier';
      }

      requestUrl += `getbytitle('${formatting.encodeQueryParameter(requestTitle)}')`;
    }
    else {
      requestUrl += `getbyuniqueid('${args.options.id}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const result = await request.get<SppModel>(requestOptions);

    if ((result as any)['odata.null'] === true) {
      throw "Model not found.";
    }

    return result;
  }

  private getDocumentLibraryInfo(args: CommandArgs): Promise<ListInstance> {
    let requestUrl = `${args.options.siteUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.siteUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get<ListInstance>(requestOptions);
  }
}

export default new SppModelApplyCommand();