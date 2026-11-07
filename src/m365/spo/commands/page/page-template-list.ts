import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { PageTemplate } from './PageTemplate.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageTemplateListCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_TEMPLATE_LIST;
  }

  public get description(): string {
    return 'Lists all page templates in the given site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'FileName', 'Id', 'PageLayoutType', 'Url'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving templates...`);
    }

    try {
      const res = await odata.getAllItems<PageTemplate>(`${args.options.webUrl}/_api/sitepages/pages/templates`);
      await logger.log(res);
    }
    catch (err: any) {
      // The API returns a 404 when no templates are created on the site collection
      if (err && err.response && err.response.status && err.response.status === 404) {
        await logger.log([]);
        return;
      }

      return this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageTemplateListCommand();