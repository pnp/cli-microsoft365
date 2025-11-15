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

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  contentCenterUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('c'),
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  id: z.uuid().optional().alias('i'),
  title: z.string().optional().alias('t'),
  listTitle: z.string().optional(),
  listId: z.uuid().optional(),
  listUrl: z.string().optional(),
  viewOption: z.enum(['NewViewAsDefault', 'DoNotChangeDefault', 'TileViewAsDefault']).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SppModelApplyCommand extends SpoCommand {
  public get name(): string {
    return commands.MODEL_APPLY;
  }

  public get description(): string {
    return 'Applies (or syncs) a trained document understanding model to a document library';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.title].filter(x => x !== undefined).length === 1, {
        error: `Specify exactly one of the following options: 'id' or 'title'.`
      })
      .refine(options => [options.listTitle, options.listId, options.listUrl].filter(x => x !== undefined).length === 1, {
        error: `Specify exactly one of the following options: 'listTitle', 'listId' or 'listUrl'.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const contentCenterUrl = urlUtil.removeTrailingSlashes(args.options.contentCenterUrl);
      await spp.assertSiteIsContentCenter(contentCenterUrl, logger, this.verbose);

      let model = null;
      if (args.options.title) {
        model = await spp.getModelByTitle(contentCenterUrl, args.options.title, logger, this.verbose);
      }
      else {
        model = await spp.getModelById(contentCenterUrl, args.options.id!, logger, this.verbose);
      }

      if (this.verbose) {
        await logger.log(`Retrieving list information...`);
      }
      const listInstance = await this.getListInfo(args.options.webUrl, args.options.listId, args.options.listTitle, args.options.listUrl);

      if (listInstance.BaseType !== 1) {
        throw `The specified list is not a document library.`;
      }

      if (this.verbose) {
        await logger.log(`Applying model '${model.ModelName}' to document library '${listInstance.RootFolder.ServerRelativeUrl}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${contentCenterUrl}/_api/machinelearning/publications`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=verbose'
        },
        responseType: 'json',
        data: {
          __metadata: { type: 'Microsoft.Office.Server.ContentCenter.SPMachineLearningPublicationsEntityData' },
          Publications: {
            results: [
              {
                ModelUniqueId: model.UniqueId,
                TargetSiteUrl: args.options.webUrl,
                TargetWebServerRelativeUrl: urlUtil.getServerRelativeSiteUrl(args.options.webUrl),
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
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get<ListInstance>(requestOptions);
  }
}

export default new SppModelApplyCommand();