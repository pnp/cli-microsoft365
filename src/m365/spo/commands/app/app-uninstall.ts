import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

const appCatalogScopeOptions = { tenant: 'tenant', sitecollection: 'sitecollection' } as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .alias('i'),
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('s'),
  appCatalogScope: zod.coercedEnum(appCatalogScopeOptions).optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoAppUninstallCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_UNINSTALL;
  }

  public get description(): string {
    return 'Uninstalls an app from the site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const uninstallApp: () => Promise<void> = async (): Promise<void> => {
      const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';

      if (this.verbose) {
        await logger.logToStderr(`Uninstalling app '${args.options.id}' from the site '${args.options.siteUrl}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${formatting.encodeQueryParameter(args.options.id)}')/uninstall`,
        headers: {
          accept: 'application/json;odata=nometadata'
        }
      };

      try {
        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataPromise(err);
      }
    };

    if (args.options.force) {
      await uninstallApp();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to uninstall the app ${args.options.id} from site ${args.options.siteUrl}?` });

      if (result) {
        await uninstallApp();
      }
    }
  }
}

export default new SpoAppUninstallCommand();