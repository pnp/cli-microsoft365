import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { CanvasSection, ClientSidePage } from './clientsidepages';
import { Page } from './Page';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  section: number;
  webUrl: string;
}

class SpoPageSectionGetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_SECTION_GET;
  }

  public get description(): string {
    return 'Get information about the specified modern page section';
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
        option: '-n, --name <name>'
      },
      {
        option: '-s, --section <section>'
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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, logger, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): void => {
        const sections: CanvasSection[] = clientSidePage.sections
          .filter(section => section.order === args.options.section);

        const isJSONOutput = args.options.output === 'json';
        if (sections.length) {
          logger.log(Page.getSectionInformation(sections[0], isJSONOutput));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoPageSectionGetCommand();