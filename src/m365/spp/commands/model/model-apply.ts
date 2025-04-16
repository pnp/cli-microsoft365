import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spp } from '../../../../utils/spp.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { ListInstance } from '../../../spo/commands/list/ListInstance.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';

const options = globalOptionsZod
  .extend({
    contentCenterUrl: z.string()
      .refine(url => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      })),
    webUrl: zod.alias('u', z.string()
      .refine(url => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      }))
    ),
    id: zod.alias('i', z.string()
      .refine(id => validation.isValidGuid(id) === true, id => ({
        message: `${id} is not a valid GUID for option 'id'.`
      }))
      .optional()),
    title: zod.alias('t', z.string().optional()),
    listTitle: z.string().optional(),
    listId: z.string().refine(listId => validation.isValidGuid(listId) === true, listId => ({
      message: `${listId} is not a valid GUID for option 'listId'.`
    })).optional(),
    listUrl: z.string().optional(),
    viewOption: z.enum(['NewViewAsDefault', 'DoNotChangeDefault', 'TileViewAsDefault']).optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SppModelApplyCommand extends SpoCommand {
  public readonly viewOptions: string[] = ['NewViewAsDefault', 'DoNotChangeDefault', 'TileViewAsDefault'];

  public get name(): string {
    return commands.MODEL_APPLY;
  }

  public get description(): string {
    return 'Applies (or syncs) a trained document understanding model to a document library';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !(options.id && options.title), {
        message: `Specify either 'id' or 'title', but not both`
      })
      .refine(options => [options.listTitle, options.listId, options.listUrl].filter(x => x !== undefined).length === 1, {
        message: `Specify at least one of the following options: 'listTitle', 'listId' or 'listUrl'`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.log(`Applying a model to a document library...`);
      }

      const contentCenterUrl = urlUtil.removeTrailingSlashes(args.options.contentCenterUrl);
      await spp.assertSiteIsContentCenter(contentCenterUrl);

      const model = await spp.getModel(contentCenterUrl, args.options.title, args.options.id);
      const listInstance = await this.getListInfo(args.options.webUrl, args.options.listId, args.options.listTitle, args.options.listUrl);

      if (listInstance.BaseType !== 1) {
        throw `The specified list is not a document library.`;
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
                TargetSiteUrl: args.options.webUrl,
                TargetWebServerRelativeUrl: urlUtil.getServerRelativePath(args.options.webUrl, ''),
                TargetLibraryServerRelativeUrl: listInstance.RootFolder.ServerRelativeUrl,
                ViewOption: args.options.viewOption ?? "NewViewAsDefault"
              }
            ]
          }
        }
      };

      const result = await request.post<any>(requestOptions);
      const resultDetails = result.Details;

      if (resultDetails && resultDetails[0]?.ErrorMessage) {
        throw resultDetails[0].ErrorMessage;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getListInfo(webUrl: string, listId?: string, listTitle?: string, listUrl?: string): Promise<ListInstance> {
    let requestUrl = `${webUrl}/_api/web`;

    if (listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(listId)}')`;
    }
    else if (listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')`;
    }
    else if (listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
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