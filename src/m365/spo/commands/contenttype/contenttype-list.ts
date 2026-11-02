import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  category: z.string().optional().alias('c')
});

type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoContentTypeListCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_LIST;
  }

  public get description(): string {
    return 'Lists all available content types in the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['StringId', 'Name', 'Hidden', 'ReadOnly', 'Sealed'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let requestUrl: string = `${args.options.webUrl}/_api/web/ContentTypes?$expand=Parent`;

      if (args.options.category) {
        requestUrl += `&$filter=Group eq '${formatting.encodeQueryParameter(args.options.category as string)}'`;
      }

      const res = await odata.getAllItems<any>(requestUrl);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoContentTypeListCommand();