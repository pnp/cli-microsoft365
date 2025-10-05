import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string().optional()
    .refine(url => url === undefined || validation.isValidPowerPagesUrl(url) === true, {
      error: e => `'${e.input}' is not a valid Power Pages URL.`
    })
    .alias('u'),
  id: z.uuid().optional().alias('i'),
  name: z.string().optional().alias('n'),
  environmentName: z.string().alias('e')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpWebSiteGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.WEBSITE_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Power Pages website.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.url, options.id, options.name].filter(x => x !== undefined).length === 1, {
        error: `Specify either url, id or name, but not multiple.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving the website...`);
    }

    try {
      let item = null;

      if (args.options.id) {
        item = await powerPlatform.getWebsiteById(args.options.environmentName, args.options.id);
      }
      else if (args.options.name) {
        item = await powerPlatform.getWebsiteByName(args.options.environmentName, args.options.name);
      }
      else if (args.options.url) {
        item = await powerPlatform.getWebsiteByUrl(args.options.environmentName, args.options.url);
      }
      await logger.log(item);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpWebSiteGetCommand();