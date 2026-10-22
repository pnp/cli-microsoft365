import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import request, { CliRequestOptions } from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string().optional()
    .refine(url => url === undefined || validation.isValidPowerPagesUrl(url) === true, {
      error: e => `'${e.input}' is not a valid Power Pages URL.`
    })
    .alias('u'),
  id: z.uuid().optional().alias('i'),
  name: z.string().optional().alias('n'),
  environmentName: z.string().alias('e'),
  force: z.boolean().optional().alias('f')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpWebSiteRemoveCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.WEBSITE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Power Pages website from the list of active sites.';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.url, options.id, options.name].filter(x => x !== undefined).length === 1, {
        error: `Specify either url, id or name, but not multiple.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.verbose) {
      await logger.logToStderr(`Removing website '${args.options.id || args.options.name || args.options.url}'...`);
    }

    if (args.options.force) {
      await this.deleteWebsite(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove website '${args.options.id || args.options.name || args.options.url}'?` });

      if (result) {
        await this.deleteWebsite(args);
      }
    }
  }

  private async getWebsiteId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    if (args.options.name) {
      const website = await powerPlatform.getWebsiteByName(args.options.environmentName, args.options.name);
      return website.id;
    }

    const website = await powerPlatform.getWebsiteByUrl(args.options.environmentName, args.options.url!);
    return website.id;
  }

  private async deleteWebsite(args: CommandArgs): Promise<void> {
    try {
      const websiteId = await this.getWebsiteId(args);

      const requestOptions: CliRequestOptions = {
        url: `https://api.powerplatform.com/powerpages/environments/${args.options.environmentName}/websites/${websiteId}?api-version=2022-03-01-preview`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpWebSiteRemoveCommand();