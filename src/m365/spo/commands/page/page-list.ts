import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
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
    .alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageListCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_LIST;
  }

  public get description(): string {
    return 'Lists all modern pages in the given site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'Title'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving client-side pages...`);
      }

      let pages: any[] = [];

      const pagesList = await odata.getAllItems<any>(`${args.options.webUrl}/_api/sitepages/pages?$orderby=Title`);

      if (pagesList && pagesList.length > 0) {
        pages = pagesList;
      }

      const files = await odata.getAllItems<any>(`${args.options.webUrl}/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`);
      if (files?.length > 0) {
        const clientSidePages: any[] = files.filter(f => f.ListItemAllFields.ClientSideApplicationId === 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec');
        pages = pages.map(p => {
          const clientSidePage = clientSidePages.find(cp => cp && cp.ListItemAllFields && cp.ListItemAllFields.Id === p.Id);
          if (clientSidePage) {
            return {
              ...clientSidePage,
              ...p
            };
          }

          return p;
        });

        pages.filter(p => p.ListItemAllFields).forEach(page => delete page.ListItemAllFields.ID);
      }
      await logger.log(pages);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageListCommand();