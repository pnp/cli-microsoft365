import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().alias('u'),
  scope: z.enum(['All', 'Site', 'Web']).optional().alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoApplicationCustomizerListCommand extends SpoCommand {
  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_LIST;
  }

  public get description(): string {
    return 'Get a list of application customizers that are added to a site.';
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'Location', 'Scope', 'Id'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(args => validation.isValidSharePointUrl(args.webUrl) === true, {
      error: e => validation.isValidSharePointUrl((e.input as Options).webUrl) as string,
      path: ['webUrl']
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving application customizers...`);
    }

    const applicationCustomizers = await spo.getCustomActions(args.options.webUrl, args.options.scope, `Location eq 'ClientSideExtension.ApplicationCustomizer'`);
    await logger.log(applicationCustomizers);
  }
}

export default new SpoApplicationCustomizerListCommand();