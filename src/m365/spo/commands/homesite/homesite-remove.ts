import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class SpoHomeSiteRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_REMOVE;
  }

  public get description(): string {
    return 'Removes a Home Site';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    const removeHomeSite: () => Promise<void> = async (): Promise<void> => {
      try {
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        await this.removeHomeSiteByUrl(args.options.url, spoAdminUrl, logger);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeHomeSite();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove '${args.options.url}' as home site?` });
      if (result) {
        await removeHomeSite();
      }
    }
  }

  private async removeHomeSiteByUrl(siteUrl: string, spoAdminUrl: string, logger: Logger): Promise<void> {
    const siteAdminProperties = await spo.getSiteAdminPropertiesByUrl(siteUrl, false, logger, this.verbose);

    if (this.verbose) {
      await logger.logToStderr(`Removing '${siteUrl}' as home site...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_api/SPO.Tenant/RemoveTargetedSite`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: {
        siteId: siteAdminProperties.SiteId
      }
    };

    await request.post(requestOptions);
  }
}

export default new SpoHomeSiteRemoveCommand();