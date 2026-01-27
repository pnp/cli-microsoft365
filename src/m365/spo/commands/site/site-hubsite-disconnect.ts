import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { spo } from '../../../../utils/spo.js';

interface CommandArgs {
  options: Options;
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  asAdmin: z.boolean().optional(),
  force: z.boolean().alias('f').optional()
});
declare type Options = z.infer<typeof options>;

class SpoSiteHubSiteDisconnectCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_HUBSITE_DISCONNECT;
  }

  public get description(): string {
    return 'Disconnects the specified site collection from its hub site';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.disconnectHubSite(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to disconnect the site collection ${args.options.siteUrl} from its hub site?` });

      if (result) {
        await this.disconnectHubSite(logger, args);
      }
    }
  }

  private async disconnectHubSite(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Disconnecting site collection ${args.options.siteUrl} from its hub site...`);
      }

      const requestOptions: CliRequestOptions = {
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (!args.options.asAdmin) {
        requestOptions.url = `${args.options.siteUrl}/_api/site/JoinHubSite('00000000-0000-0000-0000-000000000000')`;
      }
      else {
        const tenantAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
        requestOptions.url = `${tenantAdminUrl}/_api/SPO.Tenant/ConnectSiteToHubSiteById`;
        requestOptions.data = {
          siteUrl: args.options.siteUrl,
          hubSiteId: '00000000-0000-0000-0000-000000000000'
        };
      }

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteHubSiteDisconnectCommand();