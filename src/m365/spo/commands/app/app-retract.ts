import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';

const appCatalogScopeOptions = { tenant: 'tenant', sitecollection: 'sitecollection' } as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .alias('i'),
  appCatalogUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .optional()
    .alias('u'),
  appCatalogScope: zod.coercedEnum(appCatalogScopeOptions).optional().alias('s'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoAppRetractCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_RETRACT;
  }

  public get description(): string {
    return 'Retracts the specified app from the specified app catalog';
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

    const retractApp = async (): Promise<void> => {
      try {
        const spoUrl = await spo.getSpoUrl(logger, this.debug);
        const appCatalogSiteUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);

        if (this.verbose) {
          await logger.logToStderr(`Retracting app...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${formatting.encodeQueryParameter(args.options.id)}')/retract`,
          headers: {
            accept: 'application/json;odata=nometadata'
          }
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataPromise(err);
      }
    };

    if (args.options.force) {
      await retractApp();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to retract the app ${args.options.id} from the app catalog?` });

      if (result) {
        await retractApp();
      }
    }
  }
}

export default new SpoAppRetractCommand();