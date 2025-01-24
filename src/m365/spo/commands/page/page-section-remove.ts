import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Page } from './Page.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  pageName: string;
  section: number;
  webUrl: string;
  force: string;
}

class SpoPageSectionRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_SECTION_REMOVE;
  }

  public get description(): string {
    return 'Remove the specified section from the modern page';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initTelemetry();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        pageName: typeof args.options.pageName !== 'undefined',
        section: typeof args.options.section !== 'undefined',
        webUrl: typeof args.options.webUrl !== 'undefined',
        force: !!args.options.force
      });
    });
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
        option: '-s, --section <section>'
      },
      {
        option: '-f, --force'
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
    if (args.options.force) {
      await this.removeSection(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove section ${args.options.section} from '${args.options.pageName}'?` });

      if (result) {
        await this.removeSection(logger, args);
      }
    }
  }

  private async removeSection(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Removing modern page section ${args.options.pageName} - ${args.options.section}...`);
      }
      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);

      const sectionToDelete = clientSidePage.sections
        .findIndex(section => section.order === args.options.section);

      if (sectionToDelete === -1) {
        throw new Error(`Section ${args.options.section} not found`);
      }

      clientSidePage.sections.splice(sectionToDelete, 1);

      const updatedContent = clientSidePage.toHtml();

      const requestOptions: any = {
        url: `${args.options
          .webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages/${args.options.pageName}')/ListItemAllFields`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue,
          'content-type': 'application/json;odata=nometadata',
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*',
          accept: 'application/json;odata=nometadata'
        },
        data: {
          CanvasContent1: updatedContent
        },
        responseType: 'json'
      };

      return request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageSectionRemoveCommand();