import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoAppInStanceListCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_INSTANCE_LIST;
  }

  public get description(): string {
    return 'Retrieves apps installed in a site';
  }

  public defaultProperties(): string[] | undefined {
    return [`Title`, `AppId`];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving installed apps in site at ${args.options.siteUrl}...`);
    }

    try {
      const apps = await odata.getAllItems<any>(`${args.options.siteUrl}/_api/web/AppTiles?$filter=AppType eq 3`);
      await logger.log(apps);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoAppInStanceListCommand();