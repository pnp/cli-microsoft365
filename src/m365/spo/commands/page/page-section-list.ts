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

class SpoPageSectionListCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_SECTION_LIST;
  }

  public get description(): string {
    return 'List sections in the specific modern page';
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);
      const sections: CanvasSection[] = clientSidePage.sections;

      const isJSONOutput = !Cli.shouldTrimOutput(args.options.output);
      if (sections.length) {
        const output = sections.map(section => Page.getSectionInformation(section, isJSONOutput));
        if (isJSONOutput) {
          await logger.log(output);
        }
        else {
          await logger.log(output.map(s => {
            return {
              order: s.order,
              columns: s.columns.length
            };
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