import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import commands from '../../commands.js';
import { AppMetadata } from './AppMetadata.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';

const appCatalogScopeOptions = { tenant: 'tenant', sitecollection: 'sitecollection' } as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional()
    .alias('i'),
  name: z.string().optional().alias('n'),
  appCatalogUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .optional()
    .alias('u'),
  appCatalogScope: zod.coercedEnum(appCatalogScopeOptions).optional().alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoAppGetCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets information about the specific app from the specified app catalog';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      })
      .refine(opts => opts.appCatalogScope !== 'sitecollection' || opts.appCatalogUrl !== undefined, {
        error: 'You must specify appCatalogUrl when the appCatalogScope is sitecollection',
        path: ['appCatalogUrl'],
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';

      const spoUrl = await spo.getSpoUrl(logger, this.debug);
      const appCatalogSiteUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);

      let appId: string;
      if (args.options.id) {
        appId = args.options.id;
      }
      else {
        if (this.verbose) {
          await logger.logToStderr(`Looking up app id for app named ${args.options.name}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${appCatalogSiteUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const app = await request.get<{ UniqueId: string }>(requestOptions);
        appId = app.UniqueId;
      }

      if (this.verbose) {
        await logger.logToStderr(`Retrieving information for app ${appId}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${formatting.encodeQueryParameter(appId)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<AppMetadata>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoAppGetCommand();