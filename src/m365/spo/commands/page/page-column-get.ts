import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { CanvasSection } from './clientsidepages.js';
import { Page } from './Page.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  pageName: z.string().alias('n'),
  section: z.number().alias('s'),
  column: z.number().alias('c')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageColumnGetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_COLUMN_GET;
  }

  public get description(): string {
    return 'Gets information about a specific column of a modern page';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);

      const sections: CanvasSection[] = clientSidePage.sections
        .filter(section => section.order === args.options.section);

      if (sections.length) {
        const isJSONOutput = !cli.shouldTrimOutput(args.options.output);
        const columns = sections[0].columns.filter(col => col.order === args.options.column);
        if (columns.length) {
          const column = Page.getColumnsInformation(columns[0], isJSONOutput);
          column.controls = columns[0].controls
            .map(control => Page.getControlsInformation(control, isJSONOutput));
          if (!isJSONOutput) {
            column.controls = (column.controls as any[])
              .map(control => `${control.id} (${control.title})`)
              .join(', ');
          }
          await logger.log(column);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageColumnGetCommand();