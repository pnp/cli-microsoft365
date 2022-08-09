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
        option: '-n, --name <name>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, logger, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): void => {
        const sections: CanvasSection[] = clientSidePage.sections;

        const isJSONOutput = args.options.output === 'json';
        if (sections.length) {
          const output = sections.map(section => Page.getSectionInformation(section, isJSONOutput));
          if (isJSONOutput) {
            logger.log(output);
          }
          else {
            logger.log(output.map(s => {
              return {
                order: s.order,
                columns: s.columns.length
              };
            }));
          }
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoPageSectionListCommand();