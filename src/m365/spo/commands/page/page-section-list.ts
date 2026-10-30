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
  pageName: z.string().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageSectionListCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_SECTION_LIST;
  }

  public get description(): string {
    return 'Lists sections in the specific modern page';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);
      const sections: CanvasSection[] = clientSidePage.sections;

      const isJSONOutput = !cli.shouldTrimOutput(args.options.output);
      if (sections.length) {
        const output = sections.map(section => Page.getSectionInformation(section, isJSONOutput));
        if (isJSONOutput) {
          await logger.log(output);
        }
        else {
          await logger.log(output.map(s => {
            const sectionOutput: any = {
              order: s.order,
              columns: s.columns.length
            };

            if (s.isVertical) {
              sectionOutput.isVertical = s.isVertical;
            }
            return sectionOutput;
          }));
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageSectionListCommand();