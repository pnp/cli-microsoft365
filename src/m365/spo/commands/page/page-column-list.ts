import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { CanvasSection } from './clientsidepages.js';
import { Page } from './Page.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  pageName: string;
  section: number;
  webUrl: string;
}

class SpoPageColumnListCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_COLUMN_LIST;
  }

  public get description(): string {
    return 'Lists columns in the specific section of a modern page';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --pageName <pageName>'
      },
      {
        option: '-s, --section <sectionId>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (isNaN(args.options.section)) {
          return `${args.options.section} is not a number`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);
      const sections: CanvasSection[] = clientSidePage.sections
        .filter(section => section.order === args.options.section);

      if (sections.length) {
        const isJSONOutput = !Cli.shouldTrimOutput(args.options.output);
        await logger.log(sections[0].columns.map(c => {
          const column = Page.getColumnsInformation(c, isJSONOutput);
          column.controls = c.controls.length;
          return column;
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageColumnListCommand();