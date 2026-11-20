import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';

const appCatalogScopeOptions = { tenant: 'tenant', sitecollection: 'sitecollection' } as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appCatalogScope: zod.coercedEnum(appCatalogScopeOptions).optional().alias('s'),
  appCatalogUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .optional()
    .alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoAppListCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists apps from the specified app catalog';
  }

  public defaultProperties(): string[] | undefined {
    return [`Title`, `ID`, `Deployed`, `AppCatalogVersion`];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.appCatalogScope !== 'sitecollection' || opts.appCatalogUrl !== undefined, {
        error: 'You must specify appCatalogUrl when the appCatalogScope is sitecollection',
        path: ['appCatalogUrl'],
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';

    try {
      const spoUrl = await spo.getSpoUrl(logger, this.debug);
      const appCatalogSiteUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving apps...`);
      }

      const apps = await odata.getAllItems<any>(`${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps`);
      await logger.log(apps);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoAppListCommand();