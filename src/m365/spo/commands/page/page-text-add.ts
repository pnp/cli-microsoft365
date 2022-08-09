import { isNumber } from 'util';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ContextInfo, spo, urlUtil, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import {
  ClientSidePage,
  ClientSideText
} from './clientsidepages';
import { Page } from './Page';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  column?: number;
  order?: number;
  pageName: string;
  section?: number;
  text: string;
  webUrl: string;
}

class SpoPageTextAddCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_TEXT_ADD;
  }

  public get description(): string {
    return 'Adds text to a modern page';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        section: typeof args.options.section !== 'undefined',
        column: typeof args.options.column !== 'undefined',
        order: typeof args.options.order !== 'undefined'
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
        option: '-t, --text <text>'
      },
      {
        option: '--section [section]'
      },
      {
        option: '--column [column]'
      },
      {
        option: '--order [order]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.section && (!isNumber(args.options.section) || args.options.section < 1)) {
          return 'The value of parameter section must be 1 or higher';
        }

        if (args.options.column && (!isNumber(args.options.column) || args.options.column < 1)) {
          return 'The value of parameter column must be 1 or higher';
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let requestDigest: string = '';

    let pageName: string = args.options.pageName;
    if (args.options.pageName.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    if (this.verbose) {
      logger.logToStderr(`Retrieving request digest...`);
    }

    spo
      .getRequestDigest(args.options.webUrl)
      .then((res: ContextInfo): Promise<ClientSidePage> => {
        // Keep the reference of request digest for subsequent requests
        requestDigest = res.FormDigestValue;

        if (this.verbose) {
          logger.logToStderr(`Retrieving modern page ${pageName}...`);
        }
        // Get Client Side Page
        return Page.getPage(pageName, args.options.webUrl, logger, this.debug, this.verbose);
      })
      .then((page: ClientSidePage): Promise<void> => {
        const section: number = (args.options.section || 1) - 1;
        const column: number = (args.options.column || 1) - 1;

        // Make sure the section is in range
        if (section >= page.sections.length) {
          throw new Error(`Invalid section '${section + 1}'`);
        }

        // Make sure the column is in range
        if (column >= page.sections[section].columns.length) {
          throw new Error(`Invalid column '${column + 1}'`);
        }

        const text: ClientSideText = new ClientSideText(args.options.text);
        if (typeof args.options.order === 'undefined') {
          page.sections[section].columns[column].addControl(text);
        }
        else {
          const order: number = args.options.order - 1;
          page.sections[section].columns[column].insertControl(text, order);
        }

        // Save the Client Side Page with updated information
        return this.saveClientSidePage(page, logger, args, pageName, requestDigest);
      })
      .then(_ => cb())
      .catch((err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private saveClientSidePage(
    clientSidePage: ClientSidePage,
    logger: Logger,
    args: CommandArgs,
    pageName: string,
    requestDigest: string
  ): Promise<void> {
    const updatedContent: string = clientSidePage.toHtml();

    if (this.debug) {
      logger.logToStderr('Updated canvas content: ');
      logger.logToStderr(updatedContent);
      logger.logToStderr('');
    }

    const requestOptions: any = {
      url: `${args.options
        .webUrl}/_api/web/getfilebyserverrelativeurl('${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages/${pageName}')/ListItemAllFields`,
      headers: {
        'X-RequestDigest': requestDigest,
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
}

module.exports = new SpoPageTextAddCommand();
